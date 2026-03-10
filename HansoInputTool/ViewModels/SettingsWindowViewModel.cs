using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using Newtonsoft.Json;

namespace HansoInputTool.ViewModels
{
    public class SettingsWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly ExcelHandler _excelHandler;
        private readonly string _ratesFilePath;
        private readonly ShortcutService _shortcutService;
        private readonly BackupService _backupService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ObservableCollection<VehicleSheetViewModel> VehicleSheetList { get; set; }

        private VehicleSheetViewModel _selectedVehicle;
        public VehicleSheetViewModel SelectedVehicle { get => _selectedVehicle; set => SetProperty(ref _selectedVehicle, value); }

        // ショートカット設定
        public ShortcutSettingsViewModel ShortcutSettingsVM { get; }

        // バックアップ設定
        private int _maxAutoBackupFiles;
        public int MaxAutoBackupFiles
        {
            get => _maxAutoBackupFiles;
            set => SetProperty(ref _maxAutoBackupFiles, Math.Max(1, Math.Min(50, value)));
        }

        private int _maxManualBackupFiles;
        public int MaxManualBackupFiles
        {
            get => _maxManualBackupFiles;
            set => SetProperty(ref _maxManualBackupFiles, Math.Max(1, Math.Min(100, value)));
        }

        // 選択中のタブインデックス
        private int _selectedTabIndex;
        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ICommand AddVehicleCommand { get; }
        public ICommand DeleteVehicleCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand ResetShortcutsCommand { get; }

        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel,
            ShortcutService shortcutService,
            BackupService backupService = null)
        {
            _excelHandler = excelHandler;
            _ratesFilePath = ratesFilePath;
            _mainViewModel = mainViewModel;
            _shortcutService = shortcutService;
            _backupService = backupService;

            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(JsonConvert.SerializeObject(currentRates));
            var currentSheets = _excelHandler.GetVehicleSheetNames();
            VehicleSheetList = new ObservableCollection<VehicleSheetViewModel>(currentSheets.Select(s => new VehicleSheetViewModel(s)));

            // ショートカット設定VMを初期化
            ShortcutSettingsVM = new ShortcutSettingsViewModel(_shortcutService.CurrentSettings);

            // バックアップ設定の初期値を読み込み
            MaxAutoBackupFiles   = _backupService?.MaxBackupFiles       ?? 10;
            MaxManualBackupFiles = _backupService?.MaxManualBackupFiles ?? 20;

            AddVehicleCommand = new RelayCommand(p => AddVehicle());
            DeleteVehicleCommand = new RelayCommand(p => DeleteVehicle(), p => SelectedVehicle != null);
            SaveCommand = new RelayCommand(p => SaveSettings(p));
            CancelCommand = new RelayCommand(p => ((Window)p).Close());
            ResetShortcutsCommand = new RelayCommand(p => ResetShortcuts());
        }

        // 旧コンストラクタ（後方互換性のため）
        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel)
            : this(currentRates, excelHandler, ratesFilePath, mainViewModel, null, null)
        {
        }

        private void AddVehicle()
        {
            var newVehicle = new VehicleSheetViewModel();
            VehicleSheetList.Add(newVehicle);
            SelectedVehicle = newVehicle;
        }

        private void DeleteVehicle()
        {
            if (SelectedVehicle == null) return;
            var sheetName = SelectedVehicle.OriginalSheetName ?? "新しい車両";
            var result = MessageBox.Show($"車両 '{sheetName}' をリストから削除しますか？\n（実際のファイルからの削除は「保存」ボタンを押した時に実行されます）", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                VehicleSheetList.Remove(SelectedVehicle);
                SelectedVehicle = null;
            }
        }

        private void ResetShortcuts()
        {
            var result = MessageBox.Show(
                "ショートカット設定をデフォルトに戻しますか？",
                "リセット確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ShortcutSettingsVM.ResetToDefaultsCommand.Execute(null);
            }
        }

        private void SaveSettings(object parameter)
        {
            // 車両シートのバリデーション
            if (VehicleSheetList.Any(v => string.IsNullOrWhiteSpace(v.VehicleTypeName)))
            {
                MessageBox.Show("車両名が空の項目があります。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var duplicate = VehicleSheetList.GroupBy(v => v.VehicleTypeName).FirstOrDefault(g => g.Count() > 1);
            if (duplicate != null)
            {
                MessageBox.Show($"車両名 '{duplicate.Key}' が重複しています。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ショートカットの重複チェック
            if (ShortcutSettingsVM.HasDuplicates(out string duplicateInfo))
            {
                var result = MessageBox.Show(
                    $"{duplicateInfo}\n\nそのまま保存しますか？",
                    "ショートカットの重複",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            try
            {
                // 車両シート設定の保存
                var originalSheetNames = _excelHandler.GetVehicleSheetNames();
                var finalSheetVMs = VehicleSheetList.ToList();

                var finalOriginalNames = finalSheetVMs.Where(vm => vm.OriginalSheetName != null).Select(vm => vm.OriginalSheetName).ToList();
                var sheetsToDelete = originalSheetNames.Except(finalOriginalNames).ToList();

                var renamedVMs = finalSheetVMs.Where(vm => vm.OriginalSheetName != null && vm.OriginalSheetName != vm.VehicleTypeName).ToList();
                var renameMap = renamedVMs.ToDictionary(vm => vm.OriginalSheetName, vm => vm.VehicleTypeName);

                var addedVMs = finalSheetVMs.Where(vm => vm.OriginalSheetName == null).ToList();
                var sheetsToAdd = new List<(string newName, string templateName)>();

                foreach (var vehicleVM in addedVMs)
                {
                    string templateSheetName = vehicleVM.Selected事業所カテゴリ == "東日本セレモニー"
                        ? "Template2"
                        : "Template1";

                    if (!_excelHandler.SheetNames.Contains(templateSheetName))
                    {
                        MessageBox.Show($"コピー元となるテンプレートシート '{templateSheetName}' が見つかりません。\nInput.xlsxに'{templateSheetName}'という名前のシートを作成してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        _excelHandler.Load();
                        return;
                    }

                    sheetsToAdd.Add((vehicleVM.VehicleTypeName, templateSheetName));
                }

                _excelHandler.SyncAllVehicleSheets(sheetsToDelete, renameMap, sheetsToAdd);
                _excelHandler.Save();

                // 料金設定の保存
                string json = JsonConvert.SerializeObject(Rates, Formatting.Indented);
                File.WriteAllText(_ratesFilePath, json);

                // ショートカット設定の保存
                if (_shortcutService != null)
                {
                    var newShortcutSettings = ShortcutSettingsVM.ToShortcutSettings();
                    _shortcutService.UpdateSettings(newShortcutSettings);
                    _shortcutService.Save();
                }

                // バックアップ保持数を反映
                if (_backupService != null)
                {
                    _backupService.MaxBackupFiles       = MaxAutoBackupFiles;
                    _backupService.MaxManualBackupFiles = MaxManualBackupFiles;
                }

                // バックアップ保持数を反映
                if (_backupService != null)
                {
                    _backupService.MaxBackupFiles       = MaxAutoBackupFiles;
                    _backupService.MaxManualBackupFiles = MaxManualBackupFiles;
                }

                _mainViewModel.UpdateRatesAndReload(Rates);

                MessageBox.Show("設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                if (parameter is Window window)
                {
                    window.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定の保存中にエラーが発生しました。\n{ex.Message}", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                _excelHandler.Load();
            }
        }
    }
}
