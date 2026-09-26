using System;
using System.Collections.Generic;
using System.Linq;
using HansoInputTool.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace HansoInputTool.Services
{
    /// <summary>
    /// DatabaseService の分割定義：読み取り・フラグJSONの変換ユーティリティまわり。
    /// テーブル/接続まわりの本体は DatabaseService.cs、セッション管理は DatabaseService.Sessions.cs、
    /// 書き込みは DatabaseService.Write.cs を参照。
    /// </summary>
    public partial class DatabaseService
    {
        #region 読み取り

        /// <summary>
        /// 指定シート・現在セッションの給油記録を日付順で返す（プレビュー表示用）。
        /// </summary>
        public List<FuelRecord> GetFuelRecords(string vehicleSheetName)
        {
            var result = new List<FuelRecord>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT id, session_id, sheet_name, day, odometer_km, liters, created_at, transport_record_id
                FROM fuel_records
                WHERE sheet_name = $sheet AND session_id = $session
                ORDER BY day, id;
            ";
            cmd.Parameters.AddWithValue("$sheet",   vehicleSheetName);
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(new FuelRecord
                {
                    Id                = reader.GetInt64(0),
                    SessionId         = reader.GetInt64(1),
                    VehicleSheetName  = reader.GetString(2),
                    Day               = (int)reader.GetInt64(3),
                    OdometerKm        = reader.GetDouble(4),
                    Liters            = reader.GetDouble(5),
                    CreatedAt         = reader.GetString(6),
                    TransportRecordId = reader.IsDBNull(7) ? null : (long?)reader.GetInt64(7),
                });
            return result;
        }

        /// <summary>
        /// 現在セッションの給油記録を全車両分まとめて返す（転記処理用）。
        /// </summary>
        public List<FuelRecord> GetAllFuelRecordsForCurrentSession()
        {
            var result = new List<FuelRecord>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT id, session_id, sheet_name, day, odometer_km, liters, created_at, transport_record_id
                FROM fuel_records
                WHERE session_id = $session
                ORDER BY sheet_name, day, id;
            ";
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(new FuelRecord
                {
                    Id                = reader.GetInt64(0),
                    SessionId         = reader.GetInt64(1),
                    VehicleSheetName  = reader.GetString(2),
                    Day               = (int)reader.GetInt64(3),
                    OdometerKm        = reader.GetDouble(4),
                    Liters            = reader.GetDouble(5),
                    CreatedAt         = reader.GetString(6),
                    TransportRecordId = reader.IsDBNull(7) ? null : (long?)reader.GetInt64(7),
                });
            return result;
        }

        /// <summary>
        /// 指定シートの全レコードを RowData リストで返す（プレビュー表示用）。
        /// </summary>
        public List<RowData> GetSheetData(
            string sheetName,
            IReadOnlyList<FlagDefinition> flags)
        {
            var result = new List<RowData>();

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT id, day, hanso_count, yuryo_km, muryo_km,
                       shinya_fee, shinya_minutes, flags_json, row_index
                FROM transport_records
                WHERE sheet_name = $sheet AND session_id = $session
                ORDER BY day, id;
            ";
            cmd.Parameters.AddWithValue("$sheet",   sheetName);
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);

            // [給油バグ修正] 給油記録は本来「特定の搬送データ行」に紐付く(transport_record_id)。
            // 同じ日に複数の搬送行があっても、紐付いている行にだけ正しく表示できる。
            // transport_record_idが未設定（新機能追加前の古いデータ）の場合のみ、
            // 従来通り「日付」で照合し、同じ日の最初の行にだけ表示するフォールバックとする。
            var fuelByTransportId = new Dictionary<long, List<FuelRecord>>();
            var fuelByDayLegacy   = new Dictionary<int, List<FuelRecord>>();
            foreach (var fuel in GetFuelRecords(sheetName))
            {
                if (fuel.TransportRecordId.HasValue)
                {
                    if (!fuelByTransportId.TryGetValue(fuel.TransportRecordId.Value, out var list))
                        fuelByTransportId[fuel.TransportRecordId.Value] = list = new List<FuelRecord>();
                    list.Add(fuel);
                }
                else
                {
                    if (!fuelByDayLegacy.TryGetValue(fuel.Day, out var list))
                        fuelByDayLegacy[fuel.Day] = list = new List<FuelRecord>();
                    list.Add(fuel);
                }
            }

            using var reader = cmd.ExecuteReader();
            int rowIndex = 3; // ExcelのrowIndexに相当する仮番号（表示順）
            var fuelAlreadyShownForDayLegacy = new HashSet<int>();
            while (reader.Read())
            {
                var flagValues = DeserializeFlags(
                    reader.IsDBNull(7) ? null : reader.GetString(7),
                    flags);

                var row = new RowData
                {
                    DbId            = reader.GetInt64(0),
                    RowIndex        = reader.IsDBNull(8) ? rowIndex : (int)reader.GetInt64(8),
                    B_Day           = reader.IsDBNull(1) ? null : (int?)reader.GetInt64(1),
                    C_Hanso         = reader.IsDBNull(2) ? null : (int?)reader.GetInt64(2),
                    D_YuryoKm       = reader.IsDBNull(3) ? null : (int?)reader.GetDouble(3),
                    E_MuryoKm       = reader.IsDBNull(4) ? null : (int?)reader.GetDouble(4),
                    H_LateFeeOotsuki = reader.IsDBNull(5) ? null : (int?)reader.GetDouble(5),
                    K_LateMinutes   = reader.IsDBNull(6) ? null : (int?)reader.GetInt64(6),
                    FlagValues      = flagValues,
                    FlagDefinitions = flags,
                };

                // 同じ日に給油記録があればプレビュー用テキストを組み立てる（例: "⛽12,345km/40L"）。
                // まず「この行に紐付いた給油記録」を優先。無ければ日付だけで紐付いた古いデータを
                // その日の最初の行にだけフォールバック表示する。
                if (fuelByTransportId.TryGetValue(row.DbId, out var fuelsForThisRow))
                {
                    row.FuelSummaryText = string.Join(" / ",
                        fuelsForThisRow.Select(f => $"⛽{f.OdometerKm:N0}km/{f.Liters:N0}L"));
                }
                else if (row.B_Day.HasValue
                         && fuelByDayLegacy.TryGetValue(row.B_Day.Value, out var fuelsThatDay)
                         && fuelAlreadyShownForDayLegacy.Add(row.B_Day.Value))
                {
                    row.FuelSummaryText = string.Join(" / ",
                        fuelsThatDay.Select(f => $"⛽{f.OdometerKm:N0}km/{f.Liters:N0}L"));
                }

                // [深夜料金バグ修正] 車両設定を優先し、未設定時のみ「大月」判定にフォールバックする
                bool isOotsuki = VehicleSettingsService?.IsFeeMode(sheetName) ?? sheetName.Contains("大月");
                row.LateValueText = isOotsuki
                    ? row.H_LateFeeOotsuki?.ToString()
                    : row.K_LateMinutes?.ToString();

                result.Add(row);
                rowIndex++;
            }

            return result;
        }

        /// <summary>
        /// 全シート名の一覧を返す（重複なし・登録順）。
        /// </summary>
        public List<string> GetAllSheetNames()
        {
            var result = new List<string>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT sheet_name FROM transport_records WHERE session_id = $session ORDER BY sheet_name;";
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(reader.GetString(0));
            return result;
        }

        /// <summary>
        /// 指定シートに残データがあるか確認する。
        /// </summary>
        public bool HasData(string sheetName)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM transport_records WHERE sheet_name = $sheet AND session_id = $session;";
            cmd.Parameters.AddWithValue("$sheet",   sheetName);
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            return (long)cmd.ExecuteScalar() > 0;
        }

        /// <summary>
        /// 全シートにデータが残っているか確認する（月初チェック用）。
        /// </summary>
        public bool HasAnyData()
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM transport_records WHERE session_id = $session;";
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            return (long)cmd.ExecuteScalar() > 0;
        }

        /// <summary>
        /// 指定シート・指定フラグIDがON(=1)のレコード数を返す（月間集計用）。
        /// </summary>
        public int GetFlagCount(string sheetName, string flagId)
        {
            using var cmd = _connection.CreateCommand();
            // flags_json は {"koryo":1,"embalming":null} 形式
            // json_extract で値を取得して 1 かどうかを判定
            cmd.CommandText = @"
                SELECT COUNT(*)
                FROM transport_records
                WHERE sheet_name = $sheet
                  AND session_id = $session
                  AND json_extract(flags_json, '$.' || $flag) = 1;
            ";
            cmd.Parameters.AddWithValue("$sheet",   sheetName);
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            cmd.Parameters.AddWithValue("$flag",    flagId);
            return (int)(long)cmd.ExecuteScalar();
        }

        #endregion

        // ────────────────────────────────────────────
        #region ユーティリティ

        /// <summary>フラグ状態を JSON 文字列に変換（Newtonsoft.Json使用）</summary>
        private static string SerializeFlags(Dictionary<string, bool> flagStates)
        {
            if (flagStates == null || flagStates.Count == 0) return "{}";

            // ON=1, OFF=nullの形式で格納する
            var obj = new JObject();
            foreach (var kv in flagStates)
                obj[kv.Key] = kv.Value ? (JToken)1 : JValue.CreateNull();

            return obj.ToString(Formatting.None);
        }

        /// <summary>JSON 文字列をフラグ辞書に変換（Newtonsoft.Json使用）</summary>
        private static Dictionary<string, int?> DeserializeFlags(
            string json,
            IReadOnlyList<FlagDefinition> flags)
        {
            var result = new Dictionary<string, int?>();

            // 全フラグをデフォルト null で初期化
            if (flags != null)
                foreach (var f in flags)
                    result[f.Id] = null;

            if (string.IsNullOrWhiteSpace(json) || json == "{}") return result;

            // フラグIDに記号等が含まれていても安全にパースできるようNewtonsoft.Jsonを使用
            try
            {
                var obj = JObject.Parse(json);
                foreach (var prop in obj.Properties())
                {
                    var val = prop.Value.Type == JTokenType.Integer ? (int?)prop.Value.Value<int>() : null;
                    result[prop.Name] = val;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"フラグJSON解析エラー: {ex.Message} json={json}");
            }

            return result;
        }

        #endregion
    }
}
