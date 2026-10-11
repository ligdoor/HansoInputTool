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
        private static string TabSettingsFilePath => Path.Combine(BaseDataPath, "tab_settings.json");
        private static string HelpFilePath
        {
            get
            {
                const string helpFileName = "搬送実績入力ツール取扱説明書.docx";
                var dataFolderHelpPath = Path.Combine(BaseDataPath, helpFileName);
                return File.Exists(dataFolderHelpPath)
                    ? dataFolderHelpPath
                    : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", helpFileName);
            }
        }
        private static string ShortcutSettingsFilePath => Path.Combine(BaseDataPath, "shortcuts.json");

        #endregion

        #region サービス

        private readonly BackupService _backupService;
        private readonly ValidationService _validationService;
        private ShortcutService _shortcutService;
        private ExcelHandler _excelHandler;
        private DatabaseService _dbService;
        private VehicleSettingsService _vehicleSettingsService;
        private TabSettingsService _tabSettingsService;
        private SheetTabResolver _sheetTabResolver;
        private ColumnMapping _columnMap;
        private List<string> _allSheetNames;
        private FlagDefinitionService _flagService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ShortcutService ShortcutService => _shortcutService;
        public FlagDefinitionService FlagService => _flagService;
        public TabSettingsService TabSettingsService => _tabSettingsService;
        public SheetTabResolver SheetTabResolver => _sheetTabResolver;

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

        /// <summary>メイン画面の入力タブ（tab_settings.json の並び順どおり）</summary>
        public ObservableCollection<TabDefinition> InputTabs { get; } = new();

        // タブごとに最後に選んでいた車両シートを覚えておく（タブを行き来しても選択が戻る）
        private readonly Dictionary<string, string> _lastSheetByTab = new();

        private TabDefinition _selectedTab;
        public TabDefinition SelectedTab
        {
            get => _selectedTab;
            set
            {
                // タブ一覧の入れ替え中などに TabControl が一時的に null を返しても無視する
                if (value == null || ReferenceEquals(_selectedTab, value)) return;

                RememberSelectedSheet(_selectedTab);
                _selectedTab = value;
                OnPropertyChanged(nameof(SelectedTab));
                OnPropertyChanged(nameof(IsTotalTabSelected));

                if (_excelHandler == null) return;
                ApplySelectedTab(value);
            }
        }

        /// <summary>選択中のタブが月合計入力か（入力パネルの表示切替・登録処理の分岐に使う）</summary>
        public bool IsTotalTabSelected => _selectedTab?.IsTotalMode == true;

        /// <summary>起動時に、tab_settings.json の内容からタブ一覧を作る（最初のタブを選択状態にする）</summary>
        private void InitializeInputTabs()
        {
            var tabs = _tabSettingsService.Tabs;

            // 各入力方式の車両一覧は、その入力方式で最初のタブを対象にして始める
            NormalSheet.ActiveTabId = tabs.FirstOrDefault(t => !t.IsTotalMode)?.Id;
            EastSheet.ActiveTabId   = tabs.FirstOrDefault(t => t.IsTotalMode)?.Id;

            _selectedTab = tabs[0];
            InputTabs.Clear();
            foreach (var t in tabs) InputTabs.Add(t);
            OnPropertyChanged(nameof(SelectedTab));
            OnPropertyChanged(nameof(IsTotalTabSelected));
        }

        /// <summary>
        /// 設定画面での変更（タブの追加・名前変更・並び替え・削除、車両の所属変更）を、
        /// メイン画面のタブと車両一覧へ反映する。
        /// </summary>
        public void ReloadInputTabs()
        {
            var latest = _tabSettingsService.Tabs;

            // 削除されたタブを外す
            foreach (var gone in InputTabs.Where(t => latest.All(n => n.Id != t.Id)).ToList())
                InputTabs.Remove(gone);

            // 追加・名前/入力方式の更新・並び順の反映（既存のタブは同じオブジェクトを使い続ける）
            for (int i = 0; i < latest.Count; i++)
            {
                var def = latest[i];
                var existing = InputTabs.FirstOrDefault(t => t.Id == def.Id);
                if (existing == null)
                {
                    InputTabs.Insert(Math.Min(i, InputTabs.Count),
                        new TabDefinition { Id = def.Id, Name = def.Name, InputMode = def.InputMode });
                }
                else
                {
                    existing.Name = def.Name;
                    existing.InputMode = def.InputMode;
                    int current = InputTabs.IndexOf(existing);
                    if (current != i) InputTabs.Move(current, i);
                }
            }

            // 選択中のタブが無くなっていたら先頭へ
            if (_selectedTab == null || !InputTabs.Contains(_selectedTab))
                _selectedTab = InputTabs.FirstOrDefault();

            // 各入力方式の車両一覧の対象タブを、有効なタブに合わせ直す
            string ValidActive(string current, bool wantTotal)
                => InputTabs.FirstOrDefault(t => t.Id == current && t.IsTotalMode == wantTotal)?.Id
                   ?? InputTabs.FirstOrDefault(t => t.IsTotalMode == wantTotal)?.Id;

            var sel = _selectedTab;
            NormalSheet.ActiveTabId = sel != null && !sel.IsTotalMode ? sel.Id : ValidActive(NormalSheet.ActiveTabId, false);
            EastSheet.ActiveTabId   = sel != null &&  sel.IsTotalMode ? sel.Id : ValidActive(EastSheet.ActiveTabId, true);

            OnPropertyChanged(nameof(SelectedTab));
            OnPropertyChanged(nameof(IsTotalTabSelected));

            if (_excelHandler != null)
            {
                PopulateSheetCombos();
                if (!IsTotalTabSelected) UpdatePreview();
            }
        }

        private void RememberSelectedSheet(TabDefinition tab)
        {
            if (tab == null) return;
            var current = tab.IsTotalMode ? EastSheet.SelectedEastSheet : NormalSheet.SelectedNormalSheet;
            if (!string.IsNullOrEmpty(current)) _lastSheetByTab[tab.Id] = current;
        }

        /// <summary>タブが切り替わったときに、車両一覧とプレビューをそのタブに合わせる</summary>
        private void ApplySelectedTab(TabDefinition tab)
        {
            _lastSheetByTab.TryGetValue(tab.Id, out var remembered);

            if (tab.IsTotalMode)
            {
                // 対象タブが変わるときだけ車両一覧を作り直す（2タブ構成のときは従来どおり何もしない）
                if (EastSheet.ActiveTabId != tab.Id)
                {
                    EastSheet.ActiveTabId = tab.Id;
                    EastSheet.PopulateSheets(_excelHandler.GetVehicleSheetNames(), remembered ?? EastSheet.SelectedEastSheet);
                }
                // 月合計入力のタブでは、日別入力のプレビューを残さない
                PreviewData.Clear();
                SelectedRow = null;
            }
            else
            {
                if (NormalSheet.ActiveTabId != tab.Id)
                {
                    NormalSheet.ActiveTabId = tab.Id;
                    NormalSheet.PopulateSheets(_excelHandler.GetVehicleSheetNames(), remembered ?? NormalSheet.SelectedNormalSheet);
                }
                UpdatePreview();
            }
        }

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
        public ICommand OpenVehicleAnnualSummaryCommand { get; }
        public ICommand OpenAuditHistoryCommand { get; }
        public ICommand OpenPdfImportCommand { get; }
        public ICommand ClearInputDataCommand { get; }
        public ICommand SwitchSessionCommand { get; }
        public ICommand OpenLateNightCalculatorCommand { get; }

        // XAMLバインディング互換のため子VMのコマンドを公開
        public ICommand RegisterNormalCommand => NormalSheet.RegisterCommand;
        public ICommand RegisterEastCommand => EastSheet.RegisterCommand;

        #endregion

        public MainViewModel()
        {
            _backupService = new BackupService(Path.Combine(BaseDataPath, "backup_settings.json"));
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
            OnClosingCommand                 = new RelayCommand(_ => BackupDatabaseBeforeDataChange());
            OpenVehicleAnnualSummaryCommand  = new RelayCommand(_ => OpenWindow<VehicleAnnualSummaryWindow>("車両別年度集計"));
            OpenAuditHistoryCommand          = new RelayCommand(_ => OpenAuditHistory(), _ => !IsBusy);
            OpenPdfImportCommand             = new RelayCommand(_ => OpenPdfImport(),                   _ => !IsBusy);
            ClearInputDataCommand            = new RelayCommand(p => ConfirmAndClearInputData(),        p => !IsBusy);
            SwitchSessionCommand             = new RelayCommand(p => OpenSessionSwitch(),               p => !IsBusy && _dbService != null);
            OpenLateNightCalculatorCommand   = new RelayCommand(p => OpenLateNightCalculator(),          p => !IsBusy && !string.IsNullOrEmpty(NormalSheet.SelectedNormalSheet));

            PreviewDataView = CollectionViewSource.GetDefaultView(PreviewData);
        }

        public void Dispose()
        {
            if (!_disposed) { _excelHandler?.Dispose(); _dbService?.Dispose(); _disposed = true; }
        }
    }
}
