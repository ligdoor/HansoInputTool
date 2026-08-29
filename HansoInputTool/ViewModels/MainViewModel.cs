using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;
using NLog;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel : ObservableObject, IDisposable
    {
        private bool _disposed = false;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        #region 定数・パス

        private const string VersionInfoUrl = "https://raw.githubusercontent.com/ligdoor/HansoInputTool/refs/heads/master/HansoInputTool/version.json";
        private const string ReleasesPageUrl = "https://github.com/ligdoor/HansoInputTool/releases";
        private const int MaxLogLines = 200;

        private static string BaseDataPath => App.DataPath
                                                ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        private static readonly string VersionFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "version.json");
        private static string RatesFilePath => Path.Combine(BaseDataPath, "rates.json");
        private static string InputFilePath => Path.Combine(BaseDataPath, "Input.xlsx");
        private static string TemplateFilePath => Path.Combine(BaseDataPath, "Template.xlsx");
        private static string ColumnMapFilePath => Path.Combine(BaseDataPath, "column_map.json");
        private static string CustomFlagsFilePath => Path.Combine(BaseDataPath, "custom_flags.json");
        private static string DatabaseFilePath => Path.Combine(BaseDataPath, "hanso_data.db");
        private static string VehicleSettingsFilePath => Path.Combine(BaseDataPath, "vehicle_settings.json");
        private static string HelpFilePath => Path.Combine(BaseDataPath, "readme.pdf");
        private static string ShortcutSettingsFilePath => Path.Combine(BaseDataPath, "shortcuts.json");

        #endregion

        #region サービス

        private readonly BackupService _backupService;
        private readonly ValidationService _validationService;
        private ShortcutService _shortcutService;
        private ExcelHandler _excelHandler;
        private DatabaseService _dbService;
        private VehicleSettingsService _vehicleSettingsService;
        private ColumnMapping _columnMap;
        private List<string> _allSheetNames;
        private FlagDefinitionService _flagService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ShortcutService ShortcutService => _shortcutService;
        public FlagDefinitionService FlagService => _flagService;

        /// <summary>指定シートが深夜料金入力方式（料金モード）かどうかを返す（EditWindow等から利用）</summary>
        public bool IsFeeMode(string sheetName) => _excelHandler?.IsFeeMode(sheetName) ?? sheetName.Contains("大月");

        /// <summary>指定シートが給油管理表への記録対象かどうかを返す（EditWindow等から利用）</summary>
        public bool IsFuelTracked(string sheetName) => _vehicleSettingsService?.IsFuelTracked(sheetName) ?? false;

        #endregion

        #region 子ViewModel

        public NormalSheetViewModel NormalSheet { get; }
        public EastSheetViewModel EastSheet { get; }

        #endregion

        #region 共通プロパティ

        private string _appVersion = "";
        public string AppVersion { get => _appVersion; private set => SetProperty(ref _appVersion, value); }

        private string _logText;
        public string LogText { get => _logText; private set => SetProperty(ref _logText, value); }
        private readonly StringBuilder _logBuilder = new();

        private int _selectedTabIndex;
        public int SelectedTabIndex { get => _selectedTabIndex; set => SetProperty(ref _selectedTabIndex, value); }

        public ObservableCollection<RowData> PreviewData { get; } = new();
        public object PreviewDataView { get; private set; }

        private RowData _selectedRow;
        public RowData SelectedRow { get => _selectedRow; set => SetProperty(ref _selectedRow, value); }

        private string _period;
        /// <summary>
        /// 「期」の入力値。値が変わるたびappsettings.jsonに保存し、
        /// 併せて（月・R年も揃っていれば）対応するDBセッションへ自動的に切り替える。
        /// これにより「期・月・R年」の組み合わせごとにデータが区分けされ、
        /// 別の月のデータと混ざってクリア／削除されてしまうことを防ぐ。
        /// </summary>
        public string Period
        {
            get => _period;
            set
            {
                if (SetProperty(ref _period, value))
                {
                    Services.DataSetupService.SaveLastPeriodRNumber(_period, _rNumber);
                    EnsureSessionMatchesCurrentPeriod();
                }
            }
        }

        private string _month;
        /// <summary>
        /// 「月」の入力値。値が変わるたびに（期・R年も揃っていれば）対応するDBセッションへ
        /// 自動的に切り替える（Periodのコメント参照）。
        /// </summary>
        public string Month
        {
            get => _month;
            set
            {
                if (SetProperty(ref _month, value))
                    EnsureSessionMatchesCurrentPeriod();
            }
        }

        private string _rNumber;
        /// <summary>
        /// 「R年」の入力値。値が変わるたびappsettings.jsonに保存し、
        /// 併せて（期・月も揃っていれば）対応するDBセッションへ自動的に切り替える（Periodのコメント参照）。
        /// </summary>
        public string RNumber
        {
            get => _rNumber;
            set
            {
                if (SetProperty(ref _rNumber, value))
                {
                    Services.DataSetupService.SaveLastPeriodRNumber(_period, _rNumber);
                    EnsureSessionMatchesCurrentPeriod();
                }
            }
        }

        private string _eraName = "R";
        public string EraName { get => _eraName; set => SetProperty(ref _eraName, value); }

        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        #endregion

        #region コマンド

        public ICommand OpenSettingsCommand { get; }
        public ICommand OpenHelpCommand { get; }
        public ICommand CreateBackupCommand { get; }
        public ICommand RestoreBackupCommand { get; }
        public ICommand OpenBackupFolderCommand { get; }
        public ICommand EditRowCommand { get; }
        public ICommand DeleteRowCommand { get; }
        public ICommand LoadGeppoFileCommand { get; }
        public ICommand SaveInputCommand { get; }
        public ICommand TransferCommand { get; }
        public ICommand OnLoadedCommand { get; }
        public ICommand OnClosingCommand { get; }
        public ICommand OpenMonthlyReportDashboardCommand { get; }
        public ICommand OpenVehicleAnnualSummaryCommand { get; }
        public ICommand OpenPdfImportCommand { get; }
        public ICommand ClearInputDataCommand { get; }
        public ICommand SwitchSessionCommand { get; }

        // XAMLバインディング互換のため子VMのコマンドを公開
        public ICommand RegisterNormalCommand => NormalSheet.RegisterCommand;
        public ICommand RegisterEastCommand => EastSheet.RegisterCommand;

        #endregion

        public MainViewModel()
        {
            _backupService = new BackupService();
            _validationService = new ValidationService();
            NormalSheet = new NormalSheetViewModel(_validationService);
            EastSheet = new EastSheetViewModel(_validationService);

            OpenSettingsCommand              = new RelayCommand(p => OpenSettings(),                    p => !IsBusy);
            OpenHelpCommand                  = new RelayCommand(p => OpenHelp(),                        p => !IsBusy);
            CreateBackupCommand              = new RelayCommand(p => CreateManualBackup(),               p => !IsBusy);
            RestoreBackupCommand             = new RelayCommand(p => OpenRestoreBackupWindow(),          p => !IsBusy);
            OpenBackupFolderCommand          = new RelayCommand(p => _backupService.OpenBackupFolder(), p => !IsBusy);
            EditRowCommand                   = new RelayCommand(p => OpenEditWindow(),                  p => SelectedRow != null && !IsBusy);
            DeleteRowCommand                 = new RelayCommand(p => DeleteSelectedRow(),               p => SelectedRow != null && !IsBusy);
            LoadGeppoFileCommand             = new RelayCommand(p => LoadGeppoFile(),                   p => !IsBusy);
            SaveInputCommand                 = new RelayCommand(p => SaveInputFile(),                   p => !IsBusy);
            TransferCommand                  = new RelayCommand(async p => await StartTransfer(),       p => !IsBusy);
            OnLoadedCommand                  = new RelayCommand(async p => await OnWindowLoaded());
            OnClosingCommand                 = new RelayCommand(p => { });
            OpenMonthlyReportDashboardCommand = new RelayCommand(_ => OpenWindow<MonthlyReportDashboardWindow>("月報統計ダッシュボード"));
            OpenVehicleAnnualSummaryCommand  = new RelayCommand(_ => OpenWindow<VehicleAnnualSummaryWindow>("車両別年度集計"));
            OpenPdfImportCommand             = new RelayCommand(_ => OpenPdfImport(),                   _ => !IsBusy);
            ClearInputDataCommand            = new RelayCommand(p => ConfirmAndClearInputData(),        p => !IsBusy);
            SwitchSessionCommand             = new RelayCommand(p => OpenSessionSwitch(),               p => !IsBusy && _dbService != null);

            PreviewDataView = CollectionViewSource.GetDefaultView(PreviewData);
        }

        public void Dispose()
        {
            if (!_disposed) { _excelHandler?.Dispose(); _dbService?.Dispose(); _disposed = true; }
        }
    }
}
