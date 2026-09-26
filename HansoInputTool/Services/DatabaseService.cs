using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HansoInputTool.Models;
using Microsoft.Data.Sqlite;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
    public partial class DatabaseService : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly string _dbPath;
        private SqliteConnection _connection;

        /// <summary>
        /// 車両ごとの深夜入力方式（時間/料金）を判定するために使用。
        /// MainViewModel初期化時に外部から設定される（未設定の場合は「大月」判定にフォールバック）。
        /// </summary>
        public VehicleSettingsService VehicleSettingsService { get; set; }

        /// <summary>
        /// コンストラクタ。指定パスのSQLiteファイルを開き、必要なテーブルがなければ作成する。
        /// </summary>
        public DatabaseService(string dbPath)
        {
            _dbPath = dbPath;
            Open();
            EnsureTableExists();
        }

        // ────────────────────────────────────────────
        #region 初期化

        /// <summary>SQLiteデータベースファイルへの接続を開く</summary>
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
                CREATE TABLE IF NOT EXISTS fuel_records (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id          INTEGER NOT NULL DEFAULT 1,
                    sheet_name          TEXT    NOT NULL,
                    day                 INTEGER NOT NULL,
                    odometer_km         REAL    NOT NULL,
                    liters              REAL    NOT NULL,
                    created_at          TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
                    transport_record_id INTEGER
                );
                CREATE TABLE IF NOT EXISTS east_values (
                    session_id  INTEGER NOT NULL,
                    sheet_name  TEXT    NOT NULL,
                    jitsudo     REAL,
                    hanso       REAL,
                    yuryo_km    REAL,
                    muryo_km    REAL,
                    unso        REAL,
                    PRIMARY KEY (session_id, sheet_name)
                );
                CREATE INDEX IF NOT EXISTS idx_sheet_name ON transport_records(sheet_name);
                CREATE INDEX IF NOT EXISTS idx_session_id ON transport_records(session_id);
                CREATE INDEX IF NOT EXISTS idx_fuel_sheet_session ON fuel_records(sheet_name, session_id);
                CREATE INDEX IF NOT EXISTS idx_fuel_transport_record ON fuel_records(transport_record_id);
            ";
            cmd.ExecuteNonQuery();

            // 既存DBへのマイグレーション: session_id 列が無ければ追加
            MigrateAddSessionId();
            // 既存DBへのマイグレーション: is_confirmed / confirmed_at 列が無ければ追加
            MigrateAddConfirmedColumns();
            // 既存DBへのマイグレーション: is_saved / saved_at 列が無ければ追加（「保存」による月データ保護）
            MigrateAddSavedColumns();
            // 既存DBへのマイグレーション: fuel_records.transport_record_id 列が無ければ追加
            MigrateAddFuelTransportRecordId();

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

        /// <summary>既存DBにis_confirmed/confirmed_at列がなければ追加する（v1.18.0移行用・確定機能）</summary>
        private void MigrateAddConfirmedColumns()
        {
            using var check = _connection.CreateCommand();
            check.CommandText = "PRAGMA table_info(month_sessions);";
            bool hasConfirmed = false;
            using (var r = check.ExecuteReader())
                while (r.Read())
                    if (r.GetString(1) == "is_confirmed") { hasConfirmed = true; break; }

            if (!hasConfirmed)
            {
                using var alter = _connection.CreateCommand();
                alter.CommandText = @"
                    ALTER TABLE month_sessions ADD COLUMN is_confirmed INTEGER NOT NULL DEFAULT 0;
                    ALTER TABLE month_sessions ADD COLUMN confirmed_at TEXT NULL;
                ";
                alter.ExecuteNonQuery();
                Logger.Info("マイグレーション: is_confirmed / confirmed_at 列を追加しました");
            }
        }

        /// <summary>
        /// 既存DBのfuel_recordsにtransport_record_id列がなければ追加する（v1.20.2移行用）。
        /// これにより給油記録を「日付」ではなく「特定の搬送データ行」に紐付けられるようにし、
        /// 同じ日に複数の搬送行がある場合でも正しい行にだけ給油が表示されるようにする。
        /// </summary>
        private void MigrateAddFuelTransportRecordId()
        {
            using var check = _connection.CreateCommand();
            check.CommandText = "PRAGMA table_info(fuel_records);";
            bool hasColumn = false;
            using (var r = check.ExecuteReader())
                while (r.Read())
                    if (r.GetString(1) == "transport_record_id") { hasColumn = true; break; }

            if (!hasColumn)
            {
                using var alter = _connection.CreateCommand();
                alter.CommandText = "ALTER TABLE fuel_records ADD COLUMN transport_record_id INTEGER;";
                alter.ExecuteNonQuery();
                Logger.Info("マイグレーション: fuel_records.transport_record_id 列を追加しました");
            }
        }

        #endregion

        // ────────────────────────────────────────────
        #region IDisposable

        /// <summary>DB接続を閉じてリソースを解放する</summary>
        public void Dispose()
        {
            _connection?.Close();
            _connection?.Dispose();
        }

        #endregion
    }
}
