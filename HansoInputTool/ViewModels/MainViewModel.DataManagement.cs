using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using HansoInputTool.Models;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel
    {
        #region データ管理

        private void ReloadAllData()
        {
            _excelHandler.Load();
            _allSheetNames = _excelHandler.SheetNames;
            PopulateSheetCombos();
            UpdatePreview();
        }

        private void PopulateSheetCombos()
        {
            var vehicleSheets = _excelHandler.GetVehicleSheetNames();
            NormalSheet.PopulateSheets(vehicleSheets, NormalSheet.SelectedNormalSheet);
            EastSheet.PopulateSheets(vehicleSheets, EastSheet.SelectedEastSheet);
        }

        // selectRowIndex を指定すると、更新後にその行を選択状態にする。
        // DataGrid側はSelectedRowの変化を検知して自動スクロールする（MainWindow.xaml.cs参照）。
        //
        // 注意: DBモード(通常運用)ではRegisterNormalDataが返すtargetRowは実は「DBの内部ID」であり、
        // RowData.RowIndexは「表示順の仮番号」で意味が異なる。実際のDB IDはRowData.DbIdに入っている。
        // ExcelフォールバックモードではtargetRowが本物のExcel行番号でRowIndexと一致する。
        // そのためDbIdでの照合を優先し、見つからなければRowIndexでも照合する。
        private void UpdatePreview(int? selectRowIndex = null)
        {
            PreviewData.Clear();
            if (string.IsNullOrEmpty(NormalSheet.SelectedNormalSheet)) return;
            foreach (var item in _excelHandler.GetSheetDataForPreview(NormalSheet.SelectedNormalSheet))
                PreviewData.Add(item);

            if (selectRowIndex.HasValue)
            {
                SelectedRow = PreviewData.FirstOrDefault(r => r.DbId == selectRowIndex.Value)
                              ?? PreviewData.FirstOrDefault(r => r.RowIndex == selectRowIndex.Value);
            }
        }

        public void UpdateRowData(string sheetName, int rowIndex, Dictionary<string, double?> newValues, Dictionary<string, bool> flagStates)
        {
            _excelHandler.UpdateNormalData(sheetName, rowIndex, newValues, flagStates);
            UpdatePreview();
            if (_dbService == null) _excelHandler.Save();
            Log($"[{sheetName}] の {rowIndex}行目のデータを更新しました。");
        }

        /// <summary>
        /// 指定の搬送データ行（transportRecordId）に紐付く給油記録一覧を返す（編集画面用）。
        /// transport_record_idが未設定の古いデータは、日付だけで一致するものをフォールバックとして拾う。
        /// </summary>
        public List<FuelRecord> GetFuelRecordsForRow(string sheetName, long transportRecordId, int day)
        {
            var all = _dbService?.GetFuelRecords(sheetName) ?? new List<FuelRecord>();
            var linked = all.Where(f => f.TransportRecordId == transportRecordId).ToList();
            if (linked.Count > 0) return linked;
            return all.Where(f => f.TransportRecordId == null && f.Day == day).ToList();
        }

        /// <summary>給油記録を新規登録する（編集画面用）。transportRecordIdでこの搬送データ行に紐付ける。</summary>
        public void AddFuelRecord(string sheetName, long transportRecordId, int day, double odometerKm, double liters)
        {
            if (_dbService == null) throw new InvalidOperationException("給油記録の登録にはデータベースモードが必要です。");
            _dbService.InsertFuelRecord(sheetName, day, odometerKm, liters, transportRecordId);
            _excelHandler.InvalidateCache(sheetName);
            UpdatePreview();
            Log($"[{sheetName}] {day}日の給油記録を追加しました。");
        }

        /// <summary>給油記録を更新する（編集画面用。既存を削除して同じ行に紐付け直して登録し直す）</summary>
        public void UpdateFuelRecord(string sheetName, long fuelId, long transportRecordId, int day, double odometerKm, double liters)
        {
            if (_dbService == null) throw new InvalidOperationException("給油記録の更新にはデータベースモードが必要です。");
            _dbService.DeleteFuelRecord(fuelId);
            _dbService.InsertFuelRecord(sheetName, day, odometerKm, liters, transportRecordId);
            _excelHandler.InvalidateCache(sheetName);
            UpdatePreview();
            Log($"[{sheetName}] {day}日の給油記録を更新しました。");
        }

        /// <summary>給油記録を削除する（編集画面用）</summary>
        public void DeleteFuelRecord(string sheetName, long fuelId)
        {
            _dbService?.DeleteFuelRecord(fuelId);
            _excelHandler.InvalidateCache(sheetName);
            UpdatePreview();
            Log($"[{sheetName}] の給油記録を削除しました。");
        }

        public void ReloadVehicleSettings(VehicleSettings settings)
        {
            _vehicleSettingsService?.Save(settings);
            NormalSheet.RefreshFeeMode();
            Log("車両設定（深夜入力方式）を更新しました。");
        }

        public void ReloadColumnMap(ColumnMapping newMap)
        {
            _columnMap = newMap;
            _excelHandler?.UpdateColumnMap(newMap);
            Log("列マッピング設定を更新しました。");
        }

        public void UpdateRatesAndReload(Dictionary<string, RateInfo> newRates)
        {
            Rates = newRates;
            _allSheetNames = _excelHandler.SheetNames;
            PopulateSheetCombos();
            UpdatePreview();
            Log("設定が更新されました。");
        }

        public void ReloadAfterRestore()
        {
            _excelHandler.Load();
            ReloadAllData();
            Log("バックアップから復元しました。データを再読み込みしました。");
            MessageBox.Show("データを再読み込みしました。", "復元完了", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ClearInputData(bool showSuccessMessage)
        {
            // 「保存」済みの月がアクティブな場合は、そのデータには一切触れず、
            // 新規入力用の空セッションへ切り替えるだけにする（保存データを消さないため）。
            bool startedBlankSession = false;
            if (_dbService != null && _dbService.IsSessionSaved(_dbService.CurrentSessionId))
            {
                Log("--- 保存済みのデータはそのまま残し、新規入力用にクリアします ---");
                // 東日本シートの現在値（保存済み月のもの）をDBへ控えてから、新規の空セッションへ切替
                var eastValues = _excelHandler.GetAllEastValues();
                if (eastValues.Count > 0)
                    _dbService.SaveEastValues(_dbService.CurrentSessionId, eastValues);
                _dbService.StartBlankSession();
                _excelHandler.InvalidateCacheAll();
                startedBlankSession = true;
            }
            else
            {
                Log("--- 入力データをクリアします ---");
                if (_dbService != null)
                {
                    _dbService.ClearAllData();
                    _excelHandler.InvalidateCacheAll();
                    Log("[DB] 全入力データをクリアしました。");
                }
            }

            // 東日本シートはDBに保存されずExcel側にのみ値が残る。ClearData() で
            // 東日本シートのセルを確実にクリアして保存する（保存済みだった月の東日本の値は
            // 事前に SaveEastValues で控え済みのため、この時点で消えても復元できる）。
            foreach (var msg in _excelHandler.ClearData()) Log(msg);
            _excelHandler.Save();
            Log("[Excel] 東日本シートのデータをクリアしました。");

            EastSheet.ClearRegisteredSheets();

            if (startedBlankSession)
            {
                // 画面上部の期・月・R年も空にして、新しい月を入力できる状態にする
                Period = Month = RNumber = string.Empty;
            }

            UpdatePreview();
            if (showSuccessMessage)
                MessageBox.Show("入力データをクリアしました。\n（「保存」済みのデータは消えていません。「その他」→「月切替」から確認できます）",
                    "クリア完了", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ConfirmAndClearInputData()
        {
            bool isSaved = _dbService != null && _dbService.IsSessionSaved(_dbService.CurrentSessionId);

            string message = isSaved
                ? "この月のデータは「保存」済みのため消えません。\n新しい月を入力できるよう、画面をクリアします。よろしいですか？"
                : "入力中のデータをすべてクリアします。\nこの操作は元に戻せません。よろしいですか？";

            if (MessageBox.Show(message, "クリア確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                try
                {
                    ClearInputData(true);
                }
                catch (System.InvalidOperationException ex)
                {
                    // 確定済みセッションをクリアしようとした場合など、意図的にブロックしている操作
                    MessageBox.Show(ex.Message, "クリアできません", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        #endregion

        #region セッション管理

        /// <summary>
        /// 東日本シートのデータ（DBには保存されずExcelのセルにしか無い）を、セッション切替に合わせて
        /// 退避・復元する。切替前のセッションの現在値をDBへ控え、切替後のセッションに控えがあれば
        /// Excelのセルへ書き戻す。控えが無ければセルは空のままになる。
        /// </summary>
        private void SwitchEastData(long fromSessionId, long toSessionId)
        {
            if (_dbService == null || fromSessionId == toSessionId) return;
            try
            {
                var current = _excelHandler.GetAllEastValues();
                if (current.Count > 0 && _dbService.SessionExists(fromSessionId))
                    _dbService.SaveEastValues(fromSessionId, current);

                _excelHandler.ClearEastValues();
                var stored = _dbService.GetEastValues(toSessionId);
                if (stored.Count > 0)
                    _excelHandler.RestoreEastValues(stored);
                _excelHandler.Save();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "東日本シートのデータ切替中にエラーが発生しました。");
                Log("東日本シートのデータ切替でエラーが発生しました。詳細はログを確認してください。");
            }
        }

        /// <summary>
        /// 画面上部の「期・月・R年」がすべて入力された時点で、その組み合わせに対応するDBセッションへ
        /// 自動的に切り替える（存在しなければ新規作成する）。Period/Month/RNumberいずれかの値が
        /// 変わるたびに呼び出される。
        ///
        /// 【この処理を追加した理由】
        /// 以前はDBセッションが「実績月報ファイルを読み込む」機能を使った時にしか作られなかった。
        /// そのため、画面から直接データを入力するだけの通常の使い方では、常にデフォルトの
        /// セッション（session_id=1）にすべての月のデータが貯まり続けてしまい、「クリア」を実行すると
        /// 今月分だけでなく既に完了して保存しておきたかった別の月のデータまで一緒に消えてしまう
        /// 不具合があった。この処理により、期・月・R年が変わるたびに対応するセッションへ確実に
        /// 切り替わるため、クリアやDB削除の影響が「今表示している期・月・R年」のデータだけに
        /// 限定されるようになる。
        /// </summary>
        private void EnsureSessionMatchesCurrentPeriod()
        {
            if (_dbService == null) return;
            if (string.IsNullOrWhiteSpace(Period) || string.IsNullOrWhiteSpace(Month) || string.IsNullOrWhiteSpace(RNumber))
                return;

            long previousSessionId = _dbService.CurrentSessionId;
            _dbService.GetOrCreateSession(Period, Month, RNumber);

            if (_dbService.CurrentSessionId != previousSessionId && _excelHandler != null)
            {
                // セッションが切り替わった場合は、表示中のデータも切り替え後の内容に合わせて更新する
                SwitchEastData(previousSessionId, _dbService.CurrentSessionId);
                _excelHandler.InvalidateCacheAll();
                EastSheet.ClearRegisteredSheets();
                UpdatePreview();
                Log($"期・月・R年の入力に合わせてデータセッションを切替しました: {Period}期 {Month}月 {EraName}{RNumber} (session_id={_dbService.CurrentSessionId})");

                // 切替前にいたセッションが、クリア済みなどで0件になっていれば自動的に片付ける
                _dbService.CleanUpEmptySessions();
            }
        }

        /// <summary>
        /// Period/Month/RNumberの3項目をまとめて変更するときに使う。プロパティセッターを
        /// 1つずつ呼ぶと、3つが揃うまでの間（例：期だけ変わって月・R年がまだ古い値の状態）に
        /// 意図しない組み合わせで一時的な空セッションが作られてしまうことがあるため、
        /// 変更通知だけまとめて行い、最後に1回だけEnsureSessionMatchesCurrentPeriod()を呼ぶ。
        /// </summary>
        private void SetPeriodMonthRNumber(string period, string month, string rNumber)
        {
            _period  = period;
            _month   = month;
            _rNumber = rNumber;
            OnPropertyChanged(nameof(Period));
            OnPropertyChanged(nameof(Month));
            OnPropertyChanged(nameof(RNumber));
            Services.DataSetupService.SaveLastPeriodRNumber(_period, _rNumber);
            EnsureSessionMatchesCurrentPeriod();
        }

        /// <summary>
        /// 「月データの切替」ダイアログを開き、ユーザーが選んだ月にデータを切り替える。
        ///
        /// 【ダイアログを閉じた後に必ずCurrentSessionIdを確認している理由】
        /// ダイアログ内で「削除」ボタンにより今アクティブなセッションそのものを削除した場合、
        /// DatabaseService側の処理で自動的に別のセッションへ内部的に切り替わる。この場合、
        /// ユーザーは明示的に「切替」ボタンを押していないためSwitchedToSessionIdはセットされず、
        /// かつダイアログを閉じてもキャンセル扱い（DialogResult=false）になる。そのまま何もしないと
        /// 画面上の期・月・R年やプレビュー表示が、実際にアクティブなセッションとズレたままになり、
        /// 「切替ボタンが反応しない／切り替えられない」ように見えてしまう不具合があった。
        /// そのため「切替」ボタンを押した場合だけでなく、ダイアログを開く前後でCurrentSessionIdが
        /// 変化していないかを必ずチェックし、変化していれば画面表示を実際の状態に合わせて更新する。
        /// </summary>
        private void OpenSessionSwitch()
        {
            long sessionIdBeforeDialog = _dbService.CurrentSessionId;

            var vm = new SessionSwitchViewModel(_dbService);
            var win = new Views.SessionSwitchWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();

            if (vm.SwitchedToSessionId.HasValue)
            {
                // ユーザーが「切替」ボタンで明示的に選んだ場合
                _dbService.SwitchSession(vm.SwitchedToSessionId.Value);
            }
            else if (_dbService.CurrentSessionId == sessionIdBeforeDialog)
            {
                // 明示的な切替もなく、内部的な切替（削除による自動切替）も起きていない
                // → 何もせず終了（キャンセルのみ）
                return;
            }
            // else: ダイアログ内の「削除」操作により、内部的にCurrentSessionIdが変わっていた場合

            // 切替後の期・月・R年をUIに反映（3項目をまとめて設定し、余分な一時セッションを作らないようにする）
            var session = _dbService.GetAllSessions().FirstOrDefault(s => s.Id == _dbService.CurrentSessionId);
            if (session != null)
                SetPeriodMonthRNumber(session.Period, session.Month, session.RNumber);

            // 切替前にいたセッションが0件になっていれば自動的に片付ける
            _dbService.CleanUpEmptySessions();

            if (_dbService.CurrentSessionId != sessionIdBeforeDialog)
                SwitchEastData(sessionIdBeforeDialog, _dbService.CurrentSessionId);

            _excelHandler.InvalidateCacheAll();
            EastSheet.ClearRegisteredSheets();
            UpdatePreview();
            Log($"月データを切替しました: {_dbService.CurrentSessionId}");
        }

        #endregion
    }
}
