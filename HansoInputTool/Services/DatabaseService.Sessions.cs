using System;
using System.Collections.Generic;
using HansoInputTool.Models;

namespace HansoInputTool.Services
{
    /// <summary>
    /// DatabaseService の分割定義：セッション（月データ）管理まわり。
    /// テーブル/接続まわりの本体は DatabaseService.cs、書き込みは DatabaseService.Write.cs、
    /// 読み取りは DatabaseService.Read.cs を参照。
    /// </summary>
    public partial class DatabaseService
    {
        // ────────────────────────────────────────────
        #region セッション管理（複数月データ）

        /// <summary>現在のアクティブセッションID（デフォルト1）</summary>
        public long CurrentSessionId { get; private set; } = 1;

        /// <summary>既存DBにis_saved/saved_at列がなければ追加する（「保存」による月データ保護機能）</summary>
        private void MigrateAddSavedColumns()
        {
            using var check = _connection.CreateCommand();
            check.CommandText = "PRAGMA table_info(month_sessions);";
            bool hasSaved = false;
            using (var r = check.ExecuteReader())
                while (r.Read())
                    if (r.GetString(1) == "is_saved") { hasSaved = true; break; }

            if (!hasSaved)
            {
                using var alter = _connection.CreateCommand();
                alter.CommandText = @"
                    ALTER TABLE month_sessions ADD COLUMN is_saved INTEGER NOT NULL DEFAULT 0;
                    ALTER TABLE month_sessions ADD COLUMN saved_at TEXT NULL;
                ";
                alter.ExecuteNonQuery();
                Logger.Info("マイグレーション: is_saved / saved_at 列を追加しました");
            }
        }

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

            // 今のセッションが「未設定」（保存済みの月をクリアして新規入力に切り替えた直後の状態）なら、
            // 期・月・R年を入力した時点でそのセッションを引き継ぐ。期・月・R年を入力する前に
            // 打ち込んだデータが、別の「未設定」セッションに取り残されないようにするため。
            if (IsBlankSession(CurrentSessionId))
            {
                using var adopt = _connection.CreateCommand();
                adopt.CommandText = @"
                    UPDATE month_sessions
                    SET period = $p, month = $m, r_number = $r, label = $label
                    WHERE id = $id;
                ";
                adopt.Parameters.AddWithValue("$p",     period);
                adopt.Parameters.AddWithValue("$m",     month);
                adopt.Parameters.AddWithValue("$r",     rNumber);
                adopt.Parameters.AddWithValue("$label", label);
                adopt.Parameters.AddWithValue("$id",    CurrentSessionId);
                adopt.ExecuteNonQuery();
                Logger.Info($"未設定セッションを引き継ぎ: id={CurrentSessionId} label={label}");
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
                       COUNT(r.id) AS record_count, s.is_confirmed, s.confirmed_at,
                       s.is_saved, s.saved_at
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
                    IsConfirmed = reader.GetInt64(7) != 0,
                    ConfirmedAt = reader.IsDBNull(8) ? null : reader.GetString(8),
                    IsSaved     = reader.GetInt64(9) != 0,
                    SavedAt     = reader.IsDBNull(10) ? null : reader.GetString(10),
                });
            return result;
        }

        /// <summary>指定セッションが確定済みかどうかを返す</summary>
        public bool IsSessionConfirmed(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT is_confirmed FROM month_sessions WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", sessionId);
            var result = cmd.ExecuteScalar();
            return result != null && Convert.ToInt64(result) != 0;
        }

        /// <summary>
        /// セッションを「確定」状態にする。確定済みセッションのレコードは
        /// EnsureSessionEditable() のチェックにより、確定解除するまで
        /// 誤操作で編集・削除できないよう保護される。
        /// </summary>
        public void ConfirmSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE month_sessions
                SET is_confirmed = 1, confirmed_at = datetime('now','localtime')
                WHERE id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"セッション確定: id={sessionId}");
        }

        /// <summary>セッションの確定状態を解除し、再び編集できるようにする</summary>
        public void UnconfirmSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE month_sessions
                SET is_confirmed = 0, confirmed_at = NULL
                WHERE id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"セッション確定解除: id={sessionId}");
        }

        /// <summary>期・月・R年がすべて空の「未設定」セッションかどうかを返す</summary>
        private bool IsBlankSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM month_sessions WHERE id = $id AND period = '' AND month = '' AND r_number = '';";
            cmd.Parameters.AddWithValue("$id", sessionId);
            return Convert.ToInt64(cmd.ExecuteScalar()) > 0;
        }

        /// <summary>指定IDのセッションが存在するかどうかを返す</summary>
        public bool SessionExists(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM month_sessions WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", sessionId);
            return Convert.ToInt64(cmd.ExecuteScalar()) > 0;
        }

        /// <summary>
        /// 期・月・R年が空の新しい「未設定」セッションを作ってアクティブにする。
        /// 保存済みの月を「クリア」したり転記終了した後に、保存したデータには一切触れずに
        /// 画面を新規入力用の空の状態へ切り替えるために使う。
        /// 期・月・R年を入力すると、GetOrCreateSession がこのセッションを引き継ぐ。
        /// </summary>
        public long StartBlankSession()
        {
            using var ins = _connection.CreateCommand();
            ins.CommandText = @"
                INSERT INTO month_sessions (period, month, r_number, label)
                VALUES ('', '', '', '（未設定）');
                SELECT last_insert_rowid();
            ";
            CurrentSessionId = (long)ins.ExecuteScalar();
            Logger.Info($"未設定セッションを作成して切替: id={CurrentSessionId}");
            return CurrentSessionId;
        }

        /// <summary>
        /// 東日本シートの入力値（Excel側にしか無い値）をセッションごとに控えて保存する。
        /// シート名 → (項目名 → 値) の形。既に控えがあれば置き換える。
        /// </summary>
        public void SaveEastValues(long sessionId, Dictionary<string, Dictionary<string, double?>> valuesBySheet)
        {
            using var tx = _connection.BeginTransaction();

            using (var del = _connection.CreateCommand())
            {
                del.Transaction = tx;
                del.CommandText = "DELETE FROM east_values WHERE session_id = $id;";
                del.Parameters.AddWithValue("$id", sessionId);
                del.ExecuteNonQuery();
            }

            foreach (var kv in valuesBySheet)
            {
                var v = kv.Value;
                using var ins = _connection.CreateCommand();
                ins.Transaction = tx;
                ins.CommandText = @"
                    INSERT INTO east_values (session_id, sheet_name, jitsudo, hanso, yuryo_km, muryo_km, unso)
                    VALUES ($id, $sheet, $j, $h, $y, $m, $u);
                ";
                ins.Parameters.AddWithValue("$id",    sessionId);
                ins.Parameters.AddWithValue("$sheet", kv.Key);
                ins.Parameters.AddWithValue("$j", (object)v.GetValueOrDefault("延実働車輌数") ?? DBNull.Value);
                ins.Parameters.AddWithValue("$h", (object)v.GetValueOrDefault("搬送回数")   ?? DBNull.Value);
                ins.Parameters.AddWithValue("$y", (object)v.GetValueOrDefault("有料キロ数") ?? DBNull.Value);
                ins.Parameters.AddWithValue("$m", (object)v.GetValueOrDefault("無料キロ数") ?? DBNull.Value);
                ins.Parameters.AddWithValue("$u", (object)v.GetValueOrDefault("運輸実績")   ?? DBNull.Value);
                ins.ExecuteNonQuery();
            }

            tx.Commit();
        }

        /// <summary>SaveEastValues で控えた東日本シートの値を返す（控えが無ければ空）</summary>
        public Dictionary<string, Dictionary<string, double?>> GetEastValues(long sessionId)
        {
            var result = new Dictionary<string, Dictionary<string, double?>>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                SELECT sheet_name, jitsudo, hanso, yuryo_km, muryo_km, unso
                FROM east_values WHERE session_id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                double? Get(int i) => r.IsDBNull(i) ? (double?)null : r.GetDouble(i);
                result[r.GetString(0)] = new Dictionary<string, double?>
                {
                    ["延実働車輌数"] = Get(1),
                    ["搬送回数"]   = Get(2),
                    ["有料キロ数"] = Get(3),
                    ["無料キロ数"] = Get(4),
                    ["運輸実績"]   = Get(5),
                };
            }
            return result;
        }

        /// <summary>指定セッションが「保存」済み（クリアから保護されている）かどうかを返す</summary>
        public bool IsSessionSaved(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT is_saved FROM month_sessions WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", sessionId);
            var result = cmd.ExecuteScalar();
            return result != null && Convert.ToInt64(result) != 0;
        }

        /// <summary>
        /// セッションを「保存済み」にする。保存済みのセッションは ClearAllData() の対象外となり、
        /// 転記終了後の自動クリアや「クリア」でデータが消えなくなる（編集・削除は通常どおりできる）。
        /// 「確定」と違い、入力や修正はブロックしない。
        /// </summary>
        /// <returns>対象セッションが存在して更新できたらtrue</returns>
        public bool SaveSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE month_sessions
                SET is_saved = 1, saved_at = datetime('now','localtime')
                WHERE id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            bool updated = cmd.ExecuteNonQuery() > 0;
            if (updated) Logger.Info($"セッション保存（クリア保護）: id={sessionId}");
            return updated;
        }

        /// <summary>セッションの「保存済み」状態を解除し、再びクリアできるようにする</summary>
        public void UnsaveSession(long sessionId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE month_sessions
                SET is_saved = 0, saved_at = NULL
                WHERE id = $id;
            ";
            cmd.Parameters.AddWithValue("$id", sessionId);
            cmd.ExecuteNonQuery();
            Logger.Info($"セッション保存解除: id={sessionId}");
        }

        /// <summary>
        /// 指定セッションが確定済みの場合、編集操作をブロックするために例外を投げる。
        /// 登録・更新・削除・クリアの各メソッドから呼び出す共通ガード。
        /// </summary>
        private void EnsureSessionEditable(long sessionId)
        {
            if (IsSessionConfirmed(sessionId))
                throw new InvalidOperationException(
                    "このセッションは確定済みのため編集できません。編集するには先に「確定解除」してください。");
        }

        /// <summary>指定レコードidが属するsession_idを返す（存在しなければnull）</summary>
        private long? GetSessionIdForRecord(long recordId)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT session_id FROM transport_records WHERE id = $id;";
            cmd.Parameters.AddWithValue("$id", recordId);
            var result = cmd.ExecuteScalar();
            return result == null ? (long?)null : Convert.ToInt64(result);
        }

        /// <summary>指定セッションに切り替える</summary>
        public void SwitchSession(long sessionId)
        {
            CurrentSessionId = sessionId;
            Logger.Info($"セッション切替: id={sessionId}");
        }

        /// <summary>
        /// 現在アクティブなセッション（CurrentSessionId）を除いて、搬送データが1件も無い
        /// 空のセッションをまとめて削除する。「クリアして空になった月」が切替一覧に
        /// 残り続けないよう、月データの切替ダイアログを開いたときや、実際に月を
        /// 切り替えた直後に呼び出すことを想定している。今使っているセッションだけは、
        /// 仮にデータが0件でも誤って消してしまわないよう対象から除外する。
        /// </summary>
        public void CleanUpEmptySessions()
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                DELETE FROM month_sessions
                WHERE id != $current
                  AND is_saved = 0
                  AND id NOT IN (SELECT DISTINCT session_id FROM transport_records);
                DELETE FROM east_values
                WHERE session_id NOT IN (SELECT id FROM month_sessions);
            ";
            cmd.Parameters.AddWithValue("$current", CurrentSessionId);
            int affected = cmd.ExecuteNonQuery();
            if (affected > 0)
                Logger.Info($"空のセッションを{affected}件自動削除しました。");
        }

        /// <summary>指定セッションのデータをすべて削除（セッション行も削除）</summary>
        public void DeleteSession(long sessionId)
        {
            EnsureSessionEditable(sessionId);

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                DELETE FROM transport_records WHERE session_id = $id;
                DELETE FROM fuel_records      WHERE session_id = $id;
                DELETE FROM east_values       WHERE session_id = $id;
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
    }
}
