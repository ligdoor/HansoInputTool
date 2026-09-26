using System;
using System.Collections.Generic;

namespace HansoInputTool.Services
{
    /// <summary>
    /// DatabaseService の分割定義：登録・更新・削除（書き込み）まわり。
    /// テーブル/接続まわりの本体は DatabaseService.cs、セッション管理は DatabaseService.Sessions.cs、
    /// 読み取りは DatabaseService.Read.cs を参照。
    /// </summary>
    public partial class DatabaseService
    {
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
            EnsureSessionEditable(CurrentSessionId);

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
            EnsureSessionEditable(CurrentSessionId);

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
            var recordSessionId = GetSessionIdForRecord(id);
            if (recordSessionId.HasValue) EnsureSessionEditable(recordSessionId.Value);

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
            var recordSessionId = GetSessionIdForRecord(id);
            if (recordSessionId.HasValue) EnsureSessionEditable(recordSessionId.Value);

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "DELETE FROM transport_records WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", id);
            cmd.ExecuteNonQuery();
            Logger.Info($"DB削除: id={id}");
        }

        /// <summary>
        /// 月末転記後などに、現在のセッションの全データをクリアする。
        /// 「保存」済みのセッションはデータ保護のためクリアせず InvalidOperationException を投げる
        /// （転記終了後の自動クリア・「クリア」メニューから呼ばれても消えないようにするため）。
        /// 「実績月報を読み込む」のように、ユーザーが上書きを明示的に確認した上で置き換える場合だけ
        /// ignoreSaved=true を指定する。
        /// </summary>
        public void ClearAllData(bool ignoreSaved = false)
        {
            EnsureSessionEditable(CurrentSessionId);
            if (!ignoreSaved && IsSessionSaved(CurrentSessionId))
                throw new InvalidOperationException(
                    "この月のデータは「保存」済みのため、削除できません。\n" +
                    "削除するには、先に「その他」→「月切替」で「保存解除」してください。");

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                DELETE FROM transport_records WHERE session_id = $session;
                DELETE FROM fuel_records      WHERE session_id = $session;
                DELETE FROM east_values       WHERE session_id = $session;
            ";
            cmd.Parameters.AddWithValue("$session", CurrentSessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"DBクリア完了 session_id={CurrentSessionId}");
        }

        /// <summary>
        /// 給油記録を1件登録する（給油管理表への転記対象車両のみ想定）。
        /// transportRecordIdを指定すると、その搬送データ行に紐付けられ、
        /// 同じ日に複数の搬送行があってもプレビューで正しい行にのみ表示される。
        /// 確定済みセッションへの登録は EnsureSessionEditable によりブロックされる。
        /// </summary>
        public long InsertFuelRecord(string vehicleSheetName, int day, double odometerKm, double liters, long? transportRecordId = null)
        {
            EnsureSessionEditable(CurrentSessionId);

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO fuel_records (session_id, sheet_name, day, odometer_km, liters, transport_record_id)
                VALUES ($session, $sheet, $day, $km, $liters, $transportId);
                SELECT last_insert_rowid();
            ";
            cmd.Parameters.AddWithValue("$session",     CurrentSessionId);
            cmd.Parameters.AddWithValue("$sheet",       vehicleSheetName);
            cmd.Parameters.AddWithValue("$day",         day);
            cmd.Parameters.AddWithValue("$km",          odometerKm);
            cmd.Parameters.AddWithValue("$liters",      liters);
            cmd.Parameters.AddWithValue("$transportId", (object)transportRecordId ?? DBNull.Value);
            var id = (long)cmd.ExecuteScalar();
            Logger.Info($"給油記録登録: {vehicleSheetName} {day}日 {odometerKm}km {liters}L (id={id}, transport_record_id={transportRecordId})");
            return id;
        }

        /// <summary>指定の給油記録を削除する</summary>
        public void DeleteFuelRecord(long id)
        {
            using var checkCmd = _connection.CreateCommand();
            checkCmd.CommandText = "SELECT session_id FROM fuel_records WHERE id = $id;";
            checkCmd.Parameters.AddWithValue("$id", id);
            var sessionResult = checkCmd.ExecuteScalar();
            if (sessionResult != null) EnsureSessionEditable(Convert.ToInt64(sessionResult));

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "DELETE FROM fuel_records WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", id);
            cmd.ExecuteNonQuery();
            Logger.Info($"給油記録削除: id={id}");
        }

        #endregion
    }
}
