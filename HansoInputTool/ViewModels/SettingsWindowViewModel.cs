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
    public partial class SettingsWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly ExcelHandler _excelHandler;
        private readonly VehicleSettingsService _vehicleSettingsService;
        private readonly string _ratesFilePath;
        private readonly ShortcutService _shortcutService;
        private readonly BackupService _backupService;
        private readonly FlagDefinitionService _flagService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ObservableCollection<VehicleSheetViewModel> VehicleSheetList { get; set; }

        private VehicleSheetViewModel _selectedVehicle;
        public VehicleSheetViewModel SelectedVehicle
        {
            get => _selectedVehicle;
            set
            {
                if (SetProperty(ref _selectedVehicle, value))
                {
                    OnPropertyChanged(nameof(CanMoveUp));
                    OnPropertyChanged(nameof(CanMoveDown));
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        // ショートカット設定
        public ShortcutSettingsViewModel ShortcutSettingsVM { get; }

        // バックアップ設定
        private int _maxAutoBackupFiles;

        // フラグ管理
        public FlagSettingsViewModel FlagSettingsVM { get; }
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

        // 元号設定
        private string _eraName;
        public string EraName
        {
            get => _eraName;
            set => SetProperty(ref _eraName, value);
        }

        private int _eraStartYear;
        public int EraStartYear
        {
            get => _eraStartYear;
            set => SetProperty(ref _eraStartYear, value);
        }


        // ───────────────────────────────
        // タブ選択・コマンド・コンストラクタ
        // ───────────────────────────────
        // 選択中のタブインデックス
        private int _selectedTabIndex;
        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ICommand AddVehicleCommand    { get; }
        public ICommand DeleteVehicleCommand { get; }
        public ICommand MoveUpCommand        { get; }
        public ICommand MoveDownCommand      { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand ResetShortcutsCommand { get; }

        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel,
            ShortcutService shortcutService,
            BackupService backupService = null,
            FlagDefinitionService flagService = null,
            VehicleSettingsService vehicleSettingsService = null)
        {
            _excelHandler           = excelHandler;
            _ratesFilePath          = ratesFilePath;
            _mainViewModel          = mainViewModel;
            _shortcutService        = shortcutService;
            _backupService          = backupService;
            _flagService            = flagService;
            _vehicleSettingsService = vehicleSettingsService;
            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(JsonConvert.SerializeObject(currentRates));
            var currentSheets = _excelHandler.GetVehicleSheetNames();
            VehicleSheetList = new ObservableCollection<VehicleSheetViewModel>(
                currentSheets.Select(s =>
                {
                    var vm = new VehicleSheetViewModel(s);
                    if (_vehicleSettingsService != null)
                    {
                        vm.LateInputMode = _vehicleSettingsService.IsFeeMode(s) ? "fee" : "time";
                        vm.IsFuelTracked = _vehicleSettingsService.IsFuelTracked(s);
                    }
                    return vm;
                }));

            // ショートカット設定VMを初期化
            ShortcutSettingsVM = new ShortcutSettingsViewModel(_shortcutService.CurrentSettings);

            // フラグ管理VMを初期化
            FlagSettingsVM = flagService != null ? new FlagSettingsViewModel(flagService) : null;

            // バックアップ設定の初期値を読み込み
            MaxAutoBackupFiles   = _backupService?.MaxBackupFiles       ?? 10;
            MaxManualBackupFiles = _backupService?.MaxManualBackupFiles ?? 20;
            EraName = Services.DataSetupService.ReadEraNameFromSettings();
            EraStartYear = Services.DataSetupService.ReadEraStartYearFromSettings();

            // 列マッピング読み込み
            var cm = Services.DataSetupService.ReadColumnMap();
            CmDay           = cm.NormalSheet.Day;
            CmHansoCount    = cm.NormalSheet.HansoCount;
            CmYuryoKm       = cm.NormalSheet.YuryoKm;
            CmMuryoKm       = cm.NormalSheet.MuryoKm;
            CmKihonFee      = cm.NormalSheet.KihonFee;
            CmSokoFee       = cm.NormalSheet.SokoFee;
            CmShinyaFee     = cm.NormalSheet.ShinyaFee;
            CmTotalFee      = cm.NormalSheet.TotalFee;
            CmShinyaMinutes = cm.NormalSheet.ShinyaMinutes;
            CmEastJitsudo     = cm.EastSheet.Jitsudo;
            CmEastHanso       = cm.EastSheet.Hanso;
            CmEastYuryoKm     = cm.EastSheet.YuryoKm;
            CmEastMuryoKm     = cm.EastSheet.MuryoKm;
            CmEastUnsoJisseki = cm.EastSheet.UnsoJisseki;
            CmShukeiDays      = cm.ShukeiSheet.Days;
            CmShukeiHanso     = cm.ShukeiSheet.Hanso;
            CmShukeiYuryoKm   = cm.ShukeiSheet.YuryoKm;
            CmShukeiMuryoKm   = cm.ShukeiSheet.MuryoKm;
            CmShukeiTotal     = cm.ShukeiSheet.Total;

            AddVehicleCommand    = new RelayCommand(p => AddVehicle());
            DeleteVehicleCommand = new RelayCommand(p => DeleteVehicle(), p => SelectedVehicle != null);
            MoveUpCommand        = new RelayCommand(_ => MoveVehicle(-1), _ => CanMoveUp);
            MoveDownCommand      = new RelayCommand(_ => MoveVehicle(1),  _ => CanMoveDown);
            SaveCommand          = new RelayCommand(p => SaveSettings(p));
            CancelCommand        = new RelayCommand(p => ((Window)p).Close());
            ResetShortcutsCommand = new RelayCommand(p => ResetShortcuts());
        }

        // 旧コンストラクタ（後方互換性のため）
        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel)
            : this(currentRates, excelHandler, ratesFilePath, mainViewModel, null, null, null)
        {
        }

    }
}
