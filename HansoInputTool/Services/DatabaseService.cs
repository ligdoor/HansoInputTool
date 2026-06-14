using System;
using System.Collections.Generic;
using System.IO;
using HansoInputTool.Models;
using Microsoft.Data.Sqlite;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// Input.xlsx の代わりに日々の入力データを SQLite (.db) に保存・読み取りするサービス。
    /// 月末に転記処理が完了したら ClearAllData() でデータをクリアする。
    ///
    /// DB構造:
    ///   transport_records テーブル
    ///     id            INTEGER PRIMARY KEY AUTOINCREMENT
    ///     sheet_name    TEXT    -- シート名（例: 寝台車富士吉田 1）
    ///     day           INTEGER -- 日
    ///     hanso_count   INTEGER -- 搬送回数
    ///     yuryo_km      REAL    -- 有料キロ
    ///     muryo_km      REAL    -- 無料キロ
    ///     shinya_fee    REAL    -- 深夜料金（大月用）
    ///     shinya_minutes INTEGER -- 深夜時間（通常用）
    ///     flags_json    TEXT    -- フラグ状態 JSON {"koryo":1,"embalming":null} 形式
    ///     created_at    TEXT    -- 登録日時（ISO8601）
    ///     row_index     INTEGER -- 元のExcel行番号（互換用・省略可）
    /// </summary>
    public class DatabaseService : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly string _dbPath;
        private SqliteConnection _connection;

        public DatabaseService(string dbPath)
        {
            _dbPath = dbPath;
            Open();
            EnsureTableExists();
        }

        // ────────────────────────────────────────────
        #region 初期化

        private void Open()
        {
            _connection = new SqliteConnection($"Data Source={_dbPath}");
            _connection.Open();
            Logger.Info($"DB接続: {_dbPath}");
        }

        /// <summary>テーブルが存在しない場合は作成する</summary>
        private void EnsureTableExists()
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS month_sessions (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    period     TEXT    NOT NULL,
                    month      TEXT    NOT NULL,
                    r_number   TEXT    NOT NULL,
                    label      TEXT    NOT NULL,
                    created_at TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
                );
                CREATE TABLE IF NOT EXISTS transport_records (
                    id             INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id     INTEGER NOT NULL DEFAULT 1,
                    sheet_name     TEXT    NOT NULL,
                    day            INTEGER,
                    hanso_count    INTEGER,
                    yuryo_km       REAL,
                    muryo_km       REAL,
                    shinya_fee     REAL,
                    shinya_minutes INTEGER,
                    flags_json     TEXT,
                    created_at     TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
                    row_index      INTEGER
                );
                CREATE INDEX IF NOT EXISTS idx_sheet_name ON transport_records(sheet_name);
                CREATE INDEX IF NOT EXISTS idx_session_id ON transport_records(session_id);
            ";
            cmd.ExecuteNonQuery();

            // 既存DBへのマイグレーション: session_id 列が無ければ追加
            MigrateAddSessionId();

            Logger.Info("transport_records テーブル確認完了");
        }

        /// <summary>既存DBにsession_id列がなければ追加する（v1.15.0移行用）</summary>
        private void MigrateAddSessionId()
        {
            using var check = _connection.CreateCommand();
            check.CommandText = "PRAGMA table_info(transport_records);";
            bool hasSessionId = false;
            using (var r = check.ExecuteReader())
                while (r.Read())
                    if (r.GetString(1) == "session_id") { hasSessionId = true; break; }

            if (!hasSessionId)
            {
                using var alter = _connection.CreateCommand();
                alter.CommandText = "ALTER TABLE transport_records ADD COLUMN session_id INTEGER NOT NULL DEFAULT 1;";
                alter.ExecuteNonQuery();
                Logger.Info("マイグレーション: session_id 列を追加しました");
            }
        }

        #endregion

        // ────────────────────────────────────────────
        #region セッション管理（複数月データ）

        /// <summary>現在のアクティブセッションID（デフォルト1）</summary>
        public long CurrentSessionId { get; private set; } = 1;

        /// <summary>
        /// 指定した期・月・R年のセッションを取得または新規作成し、アクティブにする。
        /// </summary>
        public long GetOrCreateSession(string period, string month, string rNumber)
        {
            var label = $"{period}期 {month}月 R{rNumber}";

            // 既存セッションを検索
            using var sel = _connection.CreateCommand();
            sel.CommandText = "SELECT id FROM month_sessions WHERE period=$p AND month=$m AND r_number=$r LIMIT 1;";
            sel.Parameters.AddWithValue("$p", period);
            sel.Parameters.AddWithValue("$m", month);
            sel.Parameters.AddWithValue("$r", rNumber);
            var existing = sel.ExecuteScalar();
            if (existing != null)
            {
                CurrentSessionId = (long)existing;
                Logger.Info($"既存セッションに切替: id={CurrentSessionId} label={label}");
                return CurrentSessionId;
            }

            // 新規作成
            using var ins = _connection.CreateCommand();
            ins.CommandText = @"
                INSERT INTO month_sessions (period, month, r_number, label)
                VALUES ($p, $m, $r, $label);
                SELECT last_insert_rowid();
            ";
            ins.Parameters.AddWithValue("$p",     period);
            ins.Parameters.AddWithValue("$m",     month);
            ins.Parameters.AddWithValue("$r",     rNumber);
            ins.Parameters.AddWithValue("$label", label);
            CurrentSessionId = (long)ins.ExecuteScalar();
            Logger.Info($"新規セッション作成: id={CurrentSessionId} label={label}");
            return CurrentSessionId;
        }

        /// <summary>保存済みセッション一覧を返す（新しい順）</summary>
        public List<MonthSession> GetAllSessions()
        {
            var result = new List<MonthSession>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT s.id, s.period, s.month, s.r_number, s.label, s.created_at,
                       COUNT(r.id) AS record_count
                FROM month_sessions s
                LEFT JOIN transport_records r ON r.session_id = s.id
                GROUP BY s.id
                ORDER BY s.id DESC;
            ";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(new MonthSession
                {
                    Id          = reader.GetInt64(0),
                    Period      = reader.GetString(1),
                    Month       = reader.GetString(2),
                    RNumber     = reader.GetString(3),
                    Label       = reader.GetString(4),
                    CreatedAt   = reader.GetString(5),
                    RecordCount = (int)reader.GetInt64(6),
                });
            return result;
        }

        /// <summary>指定セッションに切り替える</summary>
        public void SwitchSession(long sessionId)
        {
            CurrentSessionId = sessionId;
            Logger.Info($"セッション切替: id={sessionId}");
        }

        /// <summary>指定セッションのデータをすべて削除（セッション行も削除）</summary>
        public void DeleteSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                DELETE FROM transport_records WHERE session_id = $id;
                DELETE FROM month_sessions     WHERE id        = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"セッション削除: id={sessionId}");

            // 削除したセッションがアクティブだった場合は残っている最新に切替
            if (CurrentSessionId == sessionId)
            {
                using var latest = _connection.CreateCommand();
                latest.CommandText = "SELECT id FROM month_sessions ORDER BY id DESC LIMIT 1;";
                var result = latest.ExecuteScalar();
                CurrentSessionId = result != null ? (long)result : 1;
                Logger.Info($"削除後セッションを切替: id={CurrentSessionId}");
            }
        }

        #endregion

        #region 書き込み（登録・更新・削除）

        /// <summary>
        /// 新規レコードを登録する。
        /// </summary>
        /// <returns>採番された id</returns>
        public long InsertRecord(
            string sheetName,
            Dictionary<string, double?> values,
            Dictionary<string, bool> flagStates,
            bool isOotsuki)
        {
            var flagsJson = SerializeFlags(flagStates);

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO transport_records
                    (session_id, sheet_name, day, hanso_count, yuryo_km, muryo_km,
                     shinya_fee, shinya_minutes, flags_json)
                VALUES
                    ($session, $sheet, $day, $hanso, $yuryo, $muryo,
                     $fee, $minutes, $flags);
                SELECT last_insert_rowid();
            ";

            double? yuryo = values.GetValueOrDefault("有料キロ(D)");
            int hanso = (yuryo.HasValue && yuryo > 0) ? 1 : 0;

            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            cmd.Parameters.AddWithValue("$sheet",   sheetName);
            cmd.Parameters.AddWithValue("$day",     (object)values.GetValueOrDefault("日(B)") ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$hanso",   hanso);
            cmd.Parameters.AddWithValue("$yuryo",   (object)yuryo ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$muryo",   (object)values.GetValueOrDefault("無料キロ(E)") ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$fee",     isOotsuki
                ? (object)(values.GetValueOrDefault("深夜料金(H)") ?? (object)DBNull.Value)
                : DBNull.Value);
            cmd.Parameters.AddWithValue("$minutes", !isOotsuki
                ? (object)(values.GetValueOrDefault("深夜時間(K)") ?? (object)DBNull.Value)
                : DBNull.Value);
            cmd.Parameters.AddWithValue("$flags",   flagsJson);

            var id = (long)cmd.ExecuteScalar();
            Logger.Info($"DB登録: id={id} sheet={sheetName} day={values.GetValueOrDefault("日(B)")}");
            return id;
        }

        /// <summary>
        /// 複数レコードを1トランザクションで一括登録する（インポート用）。
        /// 1件ずつInsertRecordを呼ぶより大幅に高速。
        /// </summary>
        public void BulkInsert(IEnumerable<(string sheetName, Dictionary<string, double?> values, Dictionary<string, bool> flagStates, bool isOotsuki)> records)
        {
            using var transaction = _connection.BeginTransaction();
            try
            {
                foreach (var (sheetName, values, flagStates, isOotsuki) in records)
                {
                    var flagsJson = SerializeFlags(flagStates);
                    using var cmd = _connection.CreateCommand();
                    cmd.Transaction = transaction;
                    cmd.CommandText = @"
                        INSERT INTO transport_records
                            (session_id, sheet_name, day, hanso_count, yuryo_km, muryo_km,
                             shinya_fee, shinya_minutes, flags_json)
                        VALUES
                            ($session, $sheet, $day, $hanso, $yuryo, $muryo,
                             $fee, $minutes, $flags);
                    ";

                    double? yuryo = values.GetValueOrDefault("有料キロ(D)");
                    int hanso = (yuryo.HasValue && yuryo > 0) ? 1 : 0;

                    cmd.Parameters.AddWithValue("$session", CurrentSessionId);
                    cmd.Parameters.AddWithValue("$sheet",   sheetName);
                    cmd.Parameters.AddWithValue("$day",     (object)values.GetValueOrDefault("日(B)") ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$hanso",   hanso);
                    cmd.Parameters.AddWithValue("$yuryo",   (object)yuryo ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$muryo",   (object)values.GetValueOrDefault("無料キロ(E)") ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("$fee",     isOotsuki
                        ? (object)(values.GetValueOrDefault("深夜料金(H)") ?? (object)DBNull.Value)
                        : DBNull.Value);
                    cmd.Parameters.AddWithValue("$minutes", !isOotsuki
                        ? (object)(values.GetValueOrDefault("深夜時間(K)") ?? (object)DBNull.Value)
                        : DBNull.Value);
                    cmd.Parameters.AddWithValue("$flags",   flagsJson);
                    cmd.ExecuteNonQuery();
                }
                transaction.Commit();
                Logger.Info($"BulkInsert完了");
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        /// <summary>
        /// 既存レコードを更新する（EditWindowからの修正用）。
        /// </summary>
        public void UpdateRecord(
            long id,
            string sheetName,
            Dictionary<string, double?> values,
            Dictionary<string, bool> flagStates,
            bool isOotsuki)
        {
            var flagsJson = SerializeFlags(flagStates);

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE transport_records SET
                    day            = $day,
                    hanso_count    = $hanso,
                    yuryo_km       = $yuryo,
                    muryo_km       = $muryo,
                    shinya_fee     = $fee,
                    shinya_minutes = $minutes,
                    flags_json     = $flags
                WHERE id = $id AND sheet_name = $sheet;
            ";

            double? yuryo = values.GetValueOrDefault("有料キロ(D)");
            int hanso = (yuryo.HasValue && yuryo > 0) ? 1 : 0;

            cmd.Parameters.AddWithValue("$id",      id);
            cmd.Parameters.AddWithValue("$sheet",   sheetName);
            cmd.Parameters.AddWithValue("$day",     (object)values.GetValueOrDefault("日(B)") ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$hanso",   hanso);
            cmd.Parameters.AddWithValue("$yuryo",   (object)yuryo ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$muryo",   (object)values.GetValueOrDefault("無料キロ(E)") ?? DBNull.Value);
            cmd.Parameters.AddWithValue("$fee",     isOotsuki
                ? (object)(values.GetValueOrDefault("深夜料金(H)") ?? (object)DBNull.Value)
                : DBNull.Value);
            cmd.Parameters.AddWithValue("$minutes", !isOotsuki
                ? (object)(values.GetValueOrDefault("深夜時間(K)") ?? (object)DBNull.Value)
                : DBNull.Value);
            cmd.Parameters.AddWithValue("$flags",   flagsJson);

            cmd.ExecuteNonQuery();
            Logger.Info($"DB更新: id={id} sheet={sheetName}");
        }

        /// <summary>
        /// 指定IDのレコードを削除する。
        /// </summary>
        public void DeleteRecord(long id)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "DELETE FROM transport_records WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", id);
            cmd.ExecuteNonQuery();
            Logger.Info($"DB削除: id={id}");
        }

        /// <summary>
        /// 月末転記後に全データをクリアする。
        /// DELETE後にAUTOINCREMENTカウンタもリセット。
        /// </summary>
        public void ClearAllData()
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "DELETE FROM transport_records WHERE session_id = $session;";
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"DBクリア完了 session_id={CurrentSessionId}");
        }

        #endregion

        // ────────────────────────────────────────────
        #region 読み取り

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

            using var reader = cmd.ExecuteReader();
            int rowIndex = 3; // ExcelのrowIndexに相当する仮番号（表示順）
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

                bool isOotsuki = sheetName.Contains("大月");
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

        /// <summary>フラグ状態を JSON 文字列に変換</summary>
        private static string SerializeFlags(Dictionary<string, bool> flagStates)
        {
            if (flagStates == null || flagStates.Count == 0) return "{}";

            var parts = new List<string>();
            foreach (var kv in flagStates)
                parts.Add($"\"{kv.Key}\":{(kv.Value ? "1" : "null")}");

            return "{" + string.Join(",", parts) + "}";
        }

        /// <summary>JSON 文字列をフラグ辞書に変換</summary>
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

            // 簡易JSONパース（Newtonsoft不使用・依存削減）
            // {"koryo":1,"embalming":null} 形式を想定
            try
            {
                var inner = json.Trim('{', '}');
                foreach (var part in inner.Split(','))
                {
                    var kv = part.Split(':');
                    if (kv.Length != 2) continue;
                    var key = kv[0].Trim().Trim('"');
                    var val = kv[1].Trim();
                    result[key] = val == "1" ? 1 : (int?)null;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"フラグJSON解析エラー: {ex.Message} json={json}");
            }

            return result;
        }

        #endregion

        // ────────────────────────────────────────────
        #region IDisposable

        public void Dispose()
        {
            _connection?.Close();
            _connection?.Dispose();
        }

        #endregion
    }
}
