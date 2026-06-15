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
