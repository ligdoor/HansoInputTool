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
            Log("--- 入力データをクリアします ---");
            if (_dbService != null)
            {
                _dbService.ClearAllData();
                _excelHandler.InvalidateCacheAll();
                Log("[DB] 全入力データをクリアしました。");

                // [No.10修正] 東日本シートはDBに保存されずExcel側にのみ値が残る。
                // DB使用時でも ClearData() を呼んで東日本シートのセルを確実にクリアし保存する。
                foreach (var msg in _excelHandler.ClearData()) Log(msg);
                _excelHandler.Save();
                Log("[Excel] 東日本シートのデータをクリアしました。");
            }
            else
            {
                foreach (var msg in _excelHandler.ClearData()) Log(msg);
                _excelHandler.Save();
            }
            EastSheet.ClearRegisteredSheets();
            UpdatePreview();
            if (showSuccessMessage)
                MessageBox.Show("入力データをクリアしました。", "クリア完了", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ConfirmAndClearInputData()
        {
            if (MessageBox.Show("入力中のデータをすべてクリアします。\nこの操作は元に戻せません。よろしいですか？",
                    "クリア確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
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

            _excelHandler.InvalidateCacheAll();
            EastSheet.ClearRegisteredSheets();
            UpdatePreview();
            Log($"月データを切替しました: {_dbService.CurrentSessionId}");
        }

        #endregion
    }
}
