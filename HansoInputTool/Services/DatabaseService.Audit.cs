using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using HansoInputTool.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace HansoInputTool.Services
{
    public partial class DatabaseService
    {
        /// <summary>Returns the newest audit events across all months.</summary>
        public List<AuditEntry> GetAuditHistory(int limit = 2000)
        {
            var result = new List<AuditEntry>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT id, event_at, user_name, action, entity_type, session_id,
                       sheet_name, record_id, day, before_json, after_json
                FROM audit_log
                ORDER BY id DESC
                LIMIT $limit;
            ";
            cmd.Parameters.AddWithValue("$limit", Math.Clamp(limit, 1, 10000));
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var entry = new AuditEntry
                {
                    Id = reader.GetInt64(0),
                    EventAt = reader.GetString(1),
                    UserName = reader.GetString(2),
                    Action = reader.GetString(3),
                    EntityType = reader.GetString(4),
                    SessionId = reader.IsDBNull(5) ? null : reader.GetInt64(5),
                    SheetName = reader.IsDBNull(6) ? "" : reader.GetString(6),
                    RecordId = reader.IsDBNull(7) ? null : reader.GetInt64(7),
                    Day = reader.IsDBNull(8) ? null : (int?)reader.GetInt64(8),
                    BeforeJson = reader.IsDBNull(9) ? "" : reader.GetString(9),
                    AfterJson = reader.IsDBNull(10) ? "" : reader.GetString(10)
                };
                entry.ChangeSummary = BuildChangeSummary(entry.BeforeJson, entry.AfterJson);
                result.Add(entry);
            }
            return result;
        }

        private void WriteAuditLog(
            string action, string entityType, string sheetName, long? recordId, int? day,
            object before, object after, long? sessionId = null,
            Microsoft.Data.Sqlite.SqliteTransaction transaction = null)
        {
            using var cmd = _connection.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = @"
                INSERT INTO audit_log
                    (event_at, user_name, action, entity_type, session_id, sheet_name,
                     record_id, day, before_json, after_json)
                VALUES
                    (datetime('now','localtime'), $user, $action, $entity, $session,
                     $sheet, $record, $day, $before, $after);
            ";
            cmd.Parameters.AddWithValue("$user", Environment.UserName ?? "");
            cmd.Parameters.AddWithValue("$action", action);
            cmd.Parameters.AddWithValue("$entity", entityType);
            cmd.Parameters.AddWithValue("$session", (object)(sessionId ?? CurrentSessionId));
            cmd.Parameters.AddWithValue("$sheet", (object)sheetName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$record", (object)recordId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$day", (object)day ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$before", before == null ? DBNull.Value : JsonConvert.SerializeObject(before));
            cmd.Parameters.AddWithValue("$after", after == null ? DBNull.Value : JsonConvert.SerializeObject(after));
            cmd.ExecuteNonQuery();
        }

        private Dictionary<string, object> ReadTransportSnapshot(long id)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT day, hanso_count, yuryo_km, muryo_km, shinya_fee, shinya_minutes, flags_json
                FROM transport_records WHERE id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", id);
            using var reader = cmd.ExecuteReader();
            if (!reader.Read()) return null;
            return new Dictionary<string, object>
            {
                ["日"] = reader.IsDBNull(0) ? null : reader.GetValue(0),
                ["搬送回数"] = reader.IsDBNull(1) ? null : reader.GetValue(1),
                ["有料キロ"] = reader.IsDBNull(2) ? null : reader.GetValue(2),
                ["無料キロ"] = reader.IsDBNull(3) ? null : reader.GetValue(3),
                ["深夜料金"] = reader.IsDBNull(4) ? null : reader.GetValue(4),
                ["深夜時間"] = reader.IsDBNull(5) ? null : reader.GetValue(5),
                ["フラグ"] = reader.IsDBNull(6) ? new JObject() : JObject.Parse(reader.GetString(6))
            };
        }

        private static Dictionary<string, object> BuildTransportSnapshot(
            Dictionary<string, double?> values, Dictionary<string, bool> flagStates, bool isFeeMode)
        {
            return new Dictionary<string, object>
            {
                ["日"] = values.GetValueOrDefault("日(B)"),
                ["搬送回数"] = values.GetValueOrDefault("搬送回数"),
                ["有料キロ"] = values.GetValueOrDefault("有料キロ(D)"),
                ["無料キロ"] = values.GetValueOrDefault("無料キロ(E)"),
                ["深夜料金"] = isFeeMode ? values.GetValueOrDefault("深夜料金(H)") : null,
                ["深夜時間"] = isFeeMode ? null : values.GetValueOrDefault("深夜時間(K)"),
                ["フラグ"] = BuildFlagsObject(flagStates)
            };
        }

        private static JObject BuildFlagsObject(Dictionary<string, bool> flagStates)
        {
            var flags = new JObject();
            if (flagStates == null) return flags;
            foreach (var flag in flagStates.OrderBy(x => x.Key))
                flags[flag.Key] = flag.Value ? new JValue(1) : JValue.CreateNull();
            return flags;
        }

        private static string BuildChangeSummary(string beforeJson, string afterJson)
        {
            try
            {
                var before = string.IsNullOrWhiteSpace(beforeJson) ? null : JObject.Parse(beforeJson);
                var after = string.IsNullOrWhiteSpace(afterJson) ? null : JObject.Parse(afterJson);
                var names = (before?.Properties().Select(p => p.Name) ?? Enumerable.Empty<string>())
                    .Union(after?.Properties().Select(p => p.Name) ?? Enumerable.Empty<string>());
                var changes = new List<string>();
                foreach (var name in names)
                {
                    var oldValue = before?[name];
                    var newValue = after?[name];
                    if (TokensAreEqual(oldValue, newValue)) continue;
                    changes.Add($"{name}: {FormatToken(oldValue)} → {FormatToken(newValue)}");
                }
                return changes.Count == 0 ? "データ内容" : string.Join(" / ", changes);
            }
            catch
            {
                return "変更内容を表示できません";
            }
        }

        private static string FormatToken(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return "（なし）";
            if (token is JArray array) return string.Join(",", array.Select(x => x.ToString()));
            return token.ToString();
        }

        private static bool TokensAreEqual(JToken left, JToken right)
        {
            if (JToken.DeepEquals(left, right)) return true;
            if (left is JValue leftValue && right is JValue rightValue &&
                (left.Type == JTokenType.Integer || left.Type == JTokenType.Float) &&
                (right.Type == JTokenType.Integer || right.Type == JTokenType.Float) &&
                decimal.TryParse(leftValue.ToString(), NumberStyles.Number, CultureInfo.InvariantCulture, out var leftNumber) &&
                decimal.TryParse(rightValue.ToString(), NumberStyles.Number, CultureInfo.InvariantCulture, out var rightNumber))
                return leftNumber == rightNumber;
            return false;
        }
    }
}
