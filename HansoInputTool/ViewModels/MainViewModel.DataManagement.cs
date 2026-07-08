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

        private void UpdatePreview()
        {
            PreviewData.Clear();
            if (string.IsNullOrEmpty(NormalSheet.SelectedNormalSheet)) return;
            foreach (var item in _excelHandler.GetSheetDataForPreview(NormalSheet.SelectedNormalSheet))
                PreviewData.Add(item);
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
                ClearInputData(true);
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
            }
        }

        private void OpenSessionSwitch()
        {
            var vm = new SessionSwitchViewModel(_dbService);
            var win = new Views.SessionSwitchWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true && vm.SwitchedToSessionId.HasValue)
            {
                _dbService.SwitchSession(vm.SwitchedToSessionId.Value);

                // 切替先の期・月・R年をUIに反映
                var session = vm.Sessions.FirstOrDefault(s => s.Id == vm.SwitchedToSessionId.Value);
                if (session != null)
                {
                    Period  = session.Period;
                    Month   = session.Month;
                    RNumber = session.RNumber;
                }

                _excelHandler.InvalidateCacheAll();
                EastSheet.ClearRegisteredSheets();
                UpdatePreview();
                Log($"月データを切替しました: {_dbService.CurrentSessionId}");
            }
        }

        #endregion
    }
}
