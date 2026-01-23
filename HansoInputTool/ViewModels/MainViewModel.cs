using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using HansoInputTool.Messaging;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;
using Microsoft.Win32;
using Newtonsoft.Json;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        #region Constants and Paths
        private const string AppName = "HansoInputTool";
        private const string CurrentVersion = "1.6.0";
        private const string GithubToken = "";
        private const string VersionInfoUrl = "https://raw.githubusercontent.com/ligdoor/HansoInputTool/refs/heads/master/version.json";
        private const string ReleasesPageUrl = "https://github.com/ligdoor/HansoInputTool/releases";

        private static readonly string BaseDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        private static readonly string RatesFilePath = Path.Combine(BaseDataPath, "rates.json");
        private static readonly string InputFilePath = Path.Combine(BaseDataPath, "Input.xlsx");
        private static readonly string TemplateFilePath = Path.Combine(BaseDataPath, "Template.xlsx");
        private static readonly string ColumnMapFilePath = Path.Combine(BaseDataPath, "column_map.json");
        private static readonly string HelpFilePath = Path.Combine(BaseDataPath, "readme.pdf");
        private static readonly string ShortcutSettingsFilePath = Path.Combine(BaseDataPath, "shortcuts.json");
        #endregion

        private readonly BackupService _backupService;
        private readonly ValidationService _validationService;
        private readonly InputValidator _inputValidator;
        private ShortcutService _shortcutService;

        #region Properties
        private ExcelHandler _excelHandler;
        public Dictionary<string, RateInfo> Rates { get; set; }
        private ColumnMapping _columnMap;
        private List<string> _allSheetNames;
        private readonly StringBuilder _logBuilder = new();
        private string _logText;
        public string LogText { get => _logText; private set => SetProperty(ref _logText, value); }
        private int _selectedTabIndex = 0;
        public int SelectedTabIndex { get => _selectedTabIndex; set => SetProperty(ref _selectedTabIndex, value); }
        public ObservableCollection<string> NormalSheets { get; } = new();
        private string _selectedNormalSheet;
        public string SelectedNormalSheet { get => _selectedNormalSheet; set { if (SetProperty(ref _selectedNormalSheet, value)) { UpdatePreview(); OnPropertyChanged(nameof(IsOotsukiSheet)); ClearNormalValidationErrors(); } } }
        public ObservableCollection<RowData> PreviewData { get; } = new();
        public ICollectionView PreviewDataView { get; }
        private RowData _selectedRow;
        public RowData SelectedRow { get => _selectedRow; set => SetProperty(ref _selectedRow, value); }
        public bool IsOotsukiSheet => SelectedNormalSheet?.Contains("大月") ?? false;

        // 通常シート入力フィールド
        private string _normalDay;
        public string NormalDay
        {
            get => _normalDay;
            set
            {
                if (SetProperty(ref _normalDay, value))
                {
                    ValidateNormalInput();
                }
            }
        }

        private string _normalYuryoKm;
        public string NormalYuryoKm
        {
            get => _normalYuryoKm;
            set
            {
                if (SetProperty(ref _normalYuryoKm, value))
                {
                    ValidateNormalInput();
                }
            }
        }

        private string _normalMuryoKm;
        public string NormalMuryoKm
        {
            get => _normalMuryoKm;
            set
            {
                if (SetProperty(ref _normalMuryoKm, value))
                {
                    ValidateNormalInput();
                }
            }
        }

        private string _normalLateValue;
        public string NormalLateValue
        {
            get => _normalLateValue;
            set
            {
                if (SetProperty(ref _normalLateValue, value))
                {
                    ValidateNormalInput();
                }
            }
        }

        private bool _isKoryo;
        public bool IsKoryo { get => _isKoryo; set => SetProperty(ref _isKoryo, value); }

        // 通常シートバリデーションエラー
        private string _normalDayError;
        public string NormalDayError { get => _normalDayError; set => SetProperty(ref _normalDayError, value); }

        private string _normalYuryoKmError;
        public string NormalYuryoKmError { get => _normalYuryoKmError; set => SetProperty(ref _normalYuryoKmError, value); }

        private string _normalMuryoKmError;
        public string NormalMuryoKmError { get => _normalMuryoKmError; set => SetProperty(ref _normalMuryoKmError, value); }

        private string _normalLateValueError;
        public string NormalLateValueError { get => _normalLateValueError; set => SetProperty(ref _normalLateValueError, value); }

        private bool _hasNormalValidationErrors;
        public bool HasNormalValidationErrors
        {
            get => _hasNormalValidationErrors;
            set
            {
                if (SetProperty(ref _hasNormalValidationErrors, value))
                {
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        // 東日本シート
        public ObservableCollection<string> EastSheets { get; } = new();
        private readonly List<string> _registeredEastSheets = new();
        private string _selectedEastSheet;
        public string SelectedEastSheet { get => _selectedEastSheet; set { if (SetProperty(ref _selectedEastSheet, value)) { UpdateEastSheetStatus(); ClearEastValidationErrors(); } } }
        private string _eastSheetStatus = "（未登録）";
        public string EastSheetStatus { get => _eastSheetStatus; set => SetProperty(ref _eastSheetStatus, value); }
        private bool _isEastSheetRegistered = false;
        public bool IsEastSheetRegistered { get => _isEastSheetRegistered; set => SetProperty(ref _isEastSheetRegistered, value); }

        private string _eastJitsudo;
        public string EastJitsudo
        {
            get => _eastJitsudo;
            set
            {
                if (SetProperty(ref _eastJitsudo, value))
                {
                    ValidateEastInput();
                }
            }
        }

        private string _eastHanso;
        public string EastHanso
        {
            get => _eastHanso;
            set
            {
                if (SetProperty(ref _eastHanso, value))
                {
                    ValidateEastInput();
                }
            }
        }

        private string _eastYuryoKm;
        public string EastYuryoKm
        {
            get => _eastYuryoKm;
            set
            {
                if (SetProperty(ref _eastYuryoKm, value))
                {
                    ValidateEastInput();
                }
            }
        }

        private string _eastMuryoKm;
        public string EastMuryoKm
        {
            get => _eastMuryoKm;
            set
            {
                if (SetProperty(ref _eastMuryoKm, value))
                {
                    ValidateEastInput();
                }
            }
        }

        private string _eastUnso;
        public string EastUnso
        {
            get => _eastUnso;
            set
            {
                if (SetProperty(ref _eastUnso, value))
                {
                    ValidateEastInput();
                }
            }
        }

        // 東日本シートバリデーションエラー
        private string _eastJitsudoError;
        public string EastJitsudoError { get => _eastJitsudoError; set => SetProperty(ref _eastJitsudoError, value); }

        private string _eastHansoError;
        public string EastHansoError { get => _eastHansoError; set => SetProperty(ref _eastHansoError, value); }

        private string _eastYuryoKmError;
        public string EastYuryoKmError { get => _eastYuryoKmError; set => SetProperty(ref _eastYuryoKmError, value); }

        private string _eastMuryoKmError;
        public string EastMuryoKmError { get => _eastMuryoKmError; set => SetProperty(ref _eastMuryoKmError, value); }

        private string _eastUnsoError;
        public string EastUnsoError { get => _eastUnsoError; set => SetProperty(ref _eastUnsoError, value); }

        private bool _hasEastValidationErrors;
        public bool HasEastValidationErrors
        {
            get => _hasEastValidationErrors;
            set
            {
                if (SetProperty(ref _hasEastValidationErrors, value))
                {
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        private string _period;
        public string Period { get => _period; set => SetProperty(ref _period, value); }
        private string _month;
        public string Month { get => _month; set => SetProperty(ref _month, value); }
        private string _rNumber;
        public string RNumber { get => _rNumber; set => SetProperty(ref _rNumber, value); }
        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        public ShortcutService ShortcutService => _shortcutService;
        #endregion

        #region Commands
        public ICommand OpenSettingsCommand { get; }
        public ICommand OpenHelpCommand { get; }
        public ICommand RegisterNormalCommand { get; }
        public ICommand RegisterEastCommand { get; }
        public ICommand EditRowCommand { get; }
        public ICommand DeleteRowCommand { get; }
        public ICommand LoadGeppoFileCommand { get; }
        public ICommand SaveInputCommand { get; }
        public ICommand TransferCommand { get; }
        public ICommand OnLoadedCommand { get; }
        public ICommand OnClosingCommand { get; }
        public ICommand OpenMonthlyReportDashboardCommand { get; }
        public ICommand CreateBackupCommand { get; }
        public ICommand RestoreBackupCommand { get; }
        public ICommand OpenBackupFolderCommand { get; }
        #endregion

        public MainViewModel()
        {
            _backupService = new BackupService();
            _validationService = new ValidationService();
            _inputValidator = new InputValidator(_validationService);
            OpenSettingsCommand = new RelayCommand(p => OpenSettings(), p => !IsBusy);
            OpenHelpCommand = new RelayCommand(p => OpenHelp(), p => !IsBusy);
            CreateBackupCommand = new RelayCommand(p => CreateManualBackup(), p => !IsBusy);
            RestoreBackupCommand = new RelayCommand(p => OpenRestoreBackupWindow(), p => !IsBusy);
            OpenBackupFolderCommand = new RelayCommand(p => _backupService.OpenBackupFolder(), p => !IsBusy);
            RegisterNormalCommand = new RelayCommand(async p => await RegisterNormal(p), p => !IsBusy && !HasNormalValidationErrors);
            RegisterEastCommand = new RelayCommand(async p => await RegisterEast(p), p => !IsBusy && !HasEastValidationErrors);
            EditRowCommand = new RelayCommand(p => OpenEditWindow(), p => SelectedRow != null && !IsBusy);
            DeleteRowCommand = new RelayCommand(p => DeleteSelectedRow(), p => SelectedRow != null && !IsBusy);
            LoadGeppoFileCommand = new RelayCommand(p => LoadGeppoFile(), p => !IsBusy);
            SaveInputCommand = new RelayCommand(p => SaveInputFile(), p => !IsBusy);
            TransferCommand = new RelayCommand(async p => await StartTransfer(), p => !IsBusy);
            OnLoadedCommand = new RelayCommand(async p => await OnWindowLoaded());
            OnClosingCommand = new RelayCommand(p => OnWindowClosing(p));
            OpenMonthlyReportDashboardCommand = new RelayCommand(_ => OpenMonthlyReportDashboard());

            PreviewDataView = CollectionViewSource.GetDefaultView(PreviewData);
        }

        #region Validation Methods

        /// <summary>
        /// 通常シートの入力値をリアルタイムバリデーション
        /// </summary>
        private void ValidateNormalInput()
        {
            var result = _inputValidator.ValidateNormalSheet(
                NormalDay,
                NormalYuryoKm,
                NormalMuryoKm,
                NormalLateValue,
                IsOotsukiSheet,
                SelectedNormalSheet);

            // 結果をプロパティに反映
            NormalDayError = result.DayError;
            NormalYuryoKmError = result.YuryoKmError;
            NormalMuryoKmError = result.MuryoKmError;
            NormalLateValueError = result.LateValueError;
            HasNormalValidationErrors = result.HasErrors;
        }

        /// <summary>
        /// 東日本シートの入力値をリアルタイムバリデーション
        /// </summary>
        private void ValidateEastInput()
        {
            var result = _inputValidator.ValidateEastSheet(
                EastJitsudo,
                EastHanso,
                EastYuryoKm,
                EastMuryoKm,
                EastUnso);

            // 結果をプロパティに反映
            EastJitsudoError = result.JitsudoError;
            EastHansoError = result.HansoError;
            EastYuryoKmError = result.YuryoKmError;
            EastMuryoKmError = result.MuryoKmError;
            EastUnsoError = result.UnsoError;
            HasEastValidationErrors = result.HasErrors;
        }

        private void ClearNormalValidationErrors()
        {
            NormalDayError = string.Empty;
            NormalYuryoKmError = string.Empty;
            NormalMuryoKmError = string.Empty;
            NormalLateValueError = string.Empty;
            HasNormalValidationErrors = false;
        }

        private void ClearEastValidationErrors()
        {
            EastJitsudoError = string.Empty;
            EastHansoError = string.Empty;
            EastYuryoKmError = string.Empty;
            EastMuryoKmError = string.Empty;
            EastUnsoError = string.Empty;
            HasEastValidationErrors = false;
        }

        #endregion

        private async Task OnWindowLoaded()
        {
            try
            {
                Logger.Info("アプリケーションの初期化を開始します。");

                if (!Directory.Exists(BaseDataPath))
                {
                    MessageBox.Show($"データフォルダが見つかりません。\n実行ファイルと同じ場所に 'data' フォルダを配置してください。", "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                    return;
                }

                Logger.Info("起動時の自動バックアップを作成します。");
                _backupService.CreateAutoBackup(InputFilePath);
                _backupService.CreateAutoBackup(TemplateFilePath);
                Log("起動時の自動バックアップを作成しました。");

                var ratesJson = await File.ReadAllTextAsync(RatesFilePath);
                Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(ratesJson);

                var columnMapJson = await File.ReadAllTextAsync(ColumnMapFilePath);
                _columnMap = JsonConvert.DeserializeObject<ColumnMapping>(columnMapJson);

                _excelHandler = new ExcelHandler(InputFilePath, TemplateFilePath, _columnMap);

                // ショートカット設定を初期化
                _shortcutService = new ShortcutService(ShortcutSettingsFilePath);
                Log("ショートカット設定を読み込みました。");

                ReloadAllData();
                await CheckForUpdate();

                if (_excelHandler.CheckRemainingData())
                {
                    var result = MessageBox.Show("前回のデータが残っています。\n全ての入力データをクリアして新規に開始しますか？", "データクリア確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        ClearInputData(true);
                    }
                }

                Logger.Info("アプリケーションの初期化が完了しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "アプリケーションの初期化中に致命的なエラーが発生しました。");
                MessageBox.Show($"アプリケーションの初期化中にエラーが発生しました。\n詳細はログファイルを確認してください。", "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private void OnWindowClosing(object parameter) { }

        #region Shortcut Processing
        /// <summary>
        /// ショートカットキーを処理する
        /// </summary>
        public bool ProcessShortcut(Key key, ModifierKeys modifiers)
        {
            if (_shortcutService == null) return false;

            var shortcuts = _shortcutService.CurrentSettings.Shortcuts;

            foreach (var kvp in shortcuts)
            {
                if (kvp.Value.Matches(key, modifiers))
                {
                    return ExecuteShortcutAction(kvp.Key);
                }
            }

            return false;
        }

        /// <summary>
        /// ショートカットアクションを実行する
        /// </summary>
        private bool ExecuteShortcutAction(string actionName)
        {
            if (IsBusy) return false;

            switch (actionName)
            {
                case "Save":
                    if (SaveInputCommand.CanExecute(null))
                    {
                        SaveInputCommand.Execute(null);
                        return true;
                    }
                    break;

                case "Register":
                    if (SelectedTabIndex == 0 && RegisterNormalCommand.CanExecute(null))
                    {
                        RegisterNormalCommand.Execute(null);
                        return true;
                    }
                    else if (SelectedTabIndex == 1 && RegisterEastCommand.CanExecute(null))
                    {
                        RegisterEastCommand.Execute(null);
                        return true;
                    }
                    break;

                case "NextSheet":
                    MoveToNextSheet();
                    return true;

                case "PrevSheet":
                    MoveToPreviousSheet();
                    return true;

                case "Transfer":
                    if (TransferCommand.CanExecute(null))
                    {
                        TransferCommand.Execute(null);
                        return true;
                    }
                    break;

                case "OpenSettings":
                    if (OpenSettingsCommand.CanExecute(null))
                    {
                        OpenSettingsCommand.Execute(null);
                        return true;
                    }
                    break;

                case "SwitchTab":
                    SelectedTabIndex = (SelectedTabIndex + 1) % 2;
                    return true;

                case "EditRow":
                    if (EditRowCommand.CanExecute(null))
                    {
                        EditRowCommand.Execute(null);
                        return true;
                    }
                    break;

                case "DeleteRow":
                    if (DeleteRowCommand.CanExecute(null))
                    {
                        DeleteRowCommand.Execute(null);
                        return true;
                    }
                    break;

                case "CreateBackup":
                    if (CreateBackupCommand.CanExecute(null))
                    {
                        CreateBackupCommand.Execute(null);
                        return true;
                    }
                    break;
            }

            return false;
        }

        /// <summary>
        /// 次のシートに移動
        /// </summary>
        private void MoveToNextSheet()
        {
            if (SelectedTabIndex == 0 && NormalSheets.Count > 0)
            {
                var currentIndex = NormalSheets.IndexOf(SelectedNormalSheet);
                if (currentIndex < NormalSheets.Count - 1)
                {
                    SelectedNormalSheet = NormalSheets[currentIndex + 1];
                }
                else
                {
                    SelectedNormalSheet = NormalSheets[0];
                }
            }
            else if (SelectedTabIndex == 1 && EastSheets.Count > 0)
            {
                var currentIndex = EastSheets.IndexOf(SelectedEastSheet);
                if (currentIndex < EastSheets.Count - 1)
                {
                    SelectedEastSheet = EastSheets[currentIndex + 1];
                }
                else
                {
                    SelectedEastSheet = EastSheets[0];
                }
            }
        }

        /// <summary>
        /// 前のシートに移動
        /// </summary>
        private void MoveToPreviousSheet()
        {
            if (SelectedTabIndex == 0 && NormalSheets.Count > 0)
            {
                var currentIndex = NormalSheets.IndexOf(SelectedNormalSheet);
                if (currentIndex > 0)
                {
                    SelectedNormalSheet = NormalSheets[currentIndex - 1];
                }
                else
                {
                    SelectedNormalSheet = NormalSheets[NormalSheets.Count - 1];
                }
            }
            else if (SelectedTabIndex == 1 && EastSheets.Count > 0)
            {
                var currentIndex = EastSheets.IndexOf(SelectedEastSheet);
                if (currentIndex > 0)
                {
                    SelectedEastSheet = EastSheets[currentIndex - 1];
                }
                else
                {
                    SelectedEastSheet = EastSheets[EastSheets.Count - 1];
                }
            }
        }
        #endregion

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
            var oldSelectedNormal = SelectedNormalSheet;
            var oldSelectedEast = SelectedEastSheet;
            NormalSheets.Clear();
            EastSheets.Clear();
            vehicleSheets.ForEach(s => { if (s.Contains("東日本")) EastSheets.Add(s); else NormalSheets.Add(s); });
            SelectedNormalSheet = NormalSheets.Contains(oldSelectedNormal) ? oldSelectedNormal : NormalSheets.FirstOrDefault();
            SelectedEastSheet = EastSheets.Contains(oldSelectedEast) ? oldSelectedEast : EastSheets.FirstOrDefault();
        }

        private void UpdatePreview()
        {
            if (string.IsNullOrEmpty(SelectedNormalSheet)) { PreviewData.Clear(); return; }
            PreviewData.Clear();
            var data = _excelHandler.GetSheetDataForPreview(SelectedNormalSheet);
            foreach (var item in data) { PreviewData.Add(item); }
        }

        public void UpdateRowData(string sheetName, int rowIndex, Dictionary<string, double?> newValues, bool isKoryo)
        {
            _excelHandler.UpdateNormalData(sheetName, rowIndex, newValues, isKoryo);
            UpdatePreview();
            _excelHandler.Save();
            Log($"[{sheetName}] の {rowIndex}行目のデータを更新しました。");
        }

        public void UpdateRatesAndReload(Dictionary<string, RateInfo> newRates)
        {
            Rates = newRates;
            _allSheetNames = _excelHandler.SheetNames;
            PopulateSheetCombos();
            UpdatePreview();
            Log("設定が更新されました。");
        }

        private void CreateManualBackup()
        {
            try
            {
                var inputBackup = _backupService.CreateManualBackup(InputFilePath, "手動保存");
                var templateBackup = _backupService.CreateManualBackup(TemplateFilePath, "手動保存");

                if (inputBackup != null && templateBackup != null)
                {
                    MessageBox.Show(
                        $"バックアップを作成しました。\n\n" +
                        $"Input.xlsx: {Path.GetFileName(inputBackup)}\n" +
                        $"Template.xlsx: {Path.GetFileName(templateBackup)}\n\n" +
                        $"保存場所: backupsフォルダ",
                        "バックアップ完了",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    Log("手動バックアップを作成しました。");
                }
                else
                {
                    MessageBox.Show("バックアップの作成に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "手動バックアップの作成中にエラーが発生しました");
                MessageBox.Show($"バックアップの作成に失敗しました。\n詳細はログファイルを確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenMonthlyReportDashboard()
        {
            try
            {
                Logger.Info("月報統計ダッシュボードを開きます");

                var dashboardWindow = new MonthlyReportDashboardWindow
                {
                    Owner = Application.Current.MainWindow
                };

                dashboardWindow.Show();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "月報統計ダッシュボードを開く際にエラーが発生");
                MessageBox.Show(
                    $"ダッシュボードを開けませんでした: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OpenRestoreBackupWindow()
        {
            var restoreVM = new RestoreBackupWindowViewModel(_backupService, InputFilePath, TemplateFilePath, this);
            var restoreWindow = new RestoreBackupWindow(restoreVM) { Owner = Application.Current.MainWindow };
            restoreWindow.ShowDialog();
        }

        public void ReloadAfterRestore()
        {
            _excelHandler.Load();
            ReloadAllData();
            Log("バックアップから復元しました。データを再読み込みしました。");
            MessageBox.Show("データを再読み込みしました。", "復元完了", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async Task RegisterNormal(object obj)
        {
            if (string.IsNullOrEmpty(SelectedNormalSheet)) { MessageBox.Show("通常シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (string.IsNullOrWhiteSpace(NormalDay)) { MessageBox.Show("日付は必須です。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            var values = new Dictionary<string, double?>();
            if (!TryParseValue(NormalDay, "日(B)", out var dayVal, silent: true)) return; values["日(B)"] = dayVal;
            if (!TryParseValue(NormalYuryoKm, "有料キロ(D)", out var yuryoKmVal, silent: true)) return;
            values["有料キロ(D)"] = yuryoKmVal.HasValue ? Math.Round(yuryoKmVal.Value, MidpointRounding.AwayFromZero) : null;
            if (!TryParseValue(NormalMuryoKm, "無料キロ(E)", out var muryoKmVal, silent: true)) return;
            values["無料キロ(E)"] = muryoKmVal.HasValue ? Math.Round(muryoKmVal.Value, MidpointRounding.AwayFromZero) : null;
            if (IsOotsukiSheet) { if (!TryParseValue(NormalLateValue, "深夜料金(H)", out var lateVal, silent: true)) return; values["深夜料金(H)"] = lateVal; }
            else { if (!TryParseValue(NormalLateValue, "深夜時間(K)", out var lateVal, silent: true)) return; values["深夜時間(K)"] = lateVal; }

            var validationResult = _validationService.ValidateNormalData(values, SelectedNormalSheet);

            if (!validationResult.IsValid)
            {
                MessageBox.Show(
                    $"入力内容にエラーがあります:\n\n{validationResult.GetErrorMessage()}",
                    "入力エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (validationResult.HasWarnings)
            {
                var result = MessageBox.Show(
                    $"以下の警告があります:\n\n{validationResult.GetWarningMessage()}\n\nそのまま登録しますか？",
                    "確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            try
            {
                var (targetRow, insertInfo) = _excelHandler.RegisterNormalData(SelectedNormalSheet, values, IsKoryo);
                UpdatePreview();
                _excelHandler.Save();
                if (!string.IsNullOrEmpty(insertInfo)) Log($"[{SelectedNormalSheet}] {insertInfo}");
                Log($"[{SelectedNormalSheet}] の {targetRow}行目にデータを登録しました。");
                NormalDay = NormalYuryoKm = NormalMuryoKm = NormalLateValue = string.Empty;
                IsKoryo = false;
                ClearNormalValidationErrors();
                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "NormalDayTextBox" });
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "通常シートへのデータ登録中にエラーが発生しました。");
                MessageBox.Show($"登録エラーが発生しました。\n詳細はログファイルを確認してください。", "登録エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task RegisterEast(object obj)
        {
            if (string.IsNullOrEmpty(SelectedEastSheet)) { MessageBox.Show("東日本シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var values = new Dictionary<string, double?>();
            if (!TryParseValue(EastJitsudo, "延実働車輌数", out var jitsudo, silent: true)) return; values["延実働車輌数"] = jitsudo;
            if (!TryParseValue(EastHanso, "搬送回数", out var hanso, silent: true)) return; values["搬送回数"] = hanso;
            if (!TryParseValue(EastYuryoKm, "有料キロ数", out var yuryo, silent: true)) return; values["有料キロ数"] = yuryo;
            if (!TryParseValue(EastMuryoKm, "無料キロ数", out var muryo, silent: true)) return; values["無料キロ数"] = muryo;
            if (!TryParseValue(EastUnso, "運輸実績", out var unso, silent: true)) return; values["運輸実績"] = unso;

            var validationResult = _validationService.ValidateEastData(values);

            if (!validationResult.IsValid)
            {
                MessageBox.Show(
                    $"入力内容にエラーがあります:\n\n{validationResult.GetErrorMessage()}",
                    "入力エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (validationResult.HasWarnings)
            {
                var result = MessageBox.Show(
                    $"以下の警告があります:\n\n{validationResult.GetWarningMessage()}\n\nそのまま登録しますか？",
                    "確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            try
            {
                _excelHandler.RegisterEastData(SelectedEastSheet, values);
                _excelHandler.Save();
                Log($"[{SelectedEastSheet}] のデータを登録しました。");
                if (!_registeredEastSheets.Contains(SelectedEastSheet)) { _registeredEastSheets.Add(SelectedEastSheet); }
                UpdateEastSheetStatus();
                EastJitsudo = EastHanso = EastYuryoKm = EastMuryoKm = EastUnso = string.Empty;
                ClearEastValidationErrors();
                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "EastJitsudoTextBox" });
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "東日本シートへのデータ登録中にエラーが発生しました。");
                MessageBox.Show($"登録エラーが発生しました。\n詳細はログファイルを確認してください。", "登録エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateEastSheetStatus()
        {
            if (string.IsNullOrEmpty(SelectedEastSheet)) { IsEastSheetRegistered = false; EastSheetStatus = ""; return; }
            if (_registeredEastSheets.Contains(SelectedEastSheet)) { IsEastSheetRegistered = true; EastSheetStatus = "✅ 登録完了"; }
            else { IsEastSheetRegistered = false; EastSheetStatus = "（未登録）"; }
        }

        private async Task StartTransfer()
        {
            if (!int.TryParse(Period, out var period) ||
                !int.TryParse(Month, out var month) ||
                !int.TryParse(RNumber, out var rNum))
            {
                MessageBox.Show("期、月、R番号を正しく入力してください。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool shouldContinue = false;

            var confirmVM = new TransferConfirmationViewModel(
                _excelHandler,
                Rates,
                _columnMap,
                Period,
                Month,
                RNumber,
                (result) => shouldContinue = result
            );

            var confirmWindow = new TransferConfirmationWindow(confirmVM)
            {
                Owner = Application.Current.MainWindow
            };

            confirmWindow.ShowDialog();

            if (!shouldContinue)
            {
                Log("転記処理がキャンセルされました。");
                return;
            }

            var dialog = new OpenFileDialog
            {
                Title = "出力先のベースフォルダを選択してください",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "フォルダを選択",
                Filter = "Folder|.",
                ValidateNames = false,
                DereferenceLinks = true
            };

            if (dialog.ShowDialog() != true)
            {
                Log("フォルダ選択がキャンセルされました。");
                return;
            }

            string outputDir = Path.GetDirectoryName(dialog.FileName);

            IsBusy = true;
            var progressVM = new ProgressWindowViewModel();
            var progressWindow = new ProgressWindow(progressVM)
            {
                Owner = Application.Current.MainWindow
            };

            var progress = new Progress<TransferProgressReport>(report =>
            {
                if (!string.IsNullOrEmpty(report.Message))
                {
                    progressVM.AppendLog(report.Message);
                }
                if (report.Total > 0)
                {
                    progressVM.UpdateProgress(report.Current, report.Total, "");
                }
            });

            progressWindow.Show();

            try
            {
                _excelHandler.Save();
                var transferService = new TransferService();
                await transferService.ExecuteAsync(
                    InputFilePath,
                    TemplateFilePath,
                    outputDir,
                    period,
                    month,
                    rNum,
                    _allSheetNames,
                    Rates,
                    _columnMap,
                    progress);

                Log("========\n転記完了\n========");
                Period = Month = RNumber = string.Empty;
                progressVM.Complete("2つのファイルの作成が完了しました。");
                ClearInputData(false);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "転記処理中にエラーが発生しました。");
                Log($"エラー: {ex.Message}");
                progressVM.ErrorComplete($"エラーが発生しました: 詳細はログファイルを確認してください。");
            }
            finally
            {
                IsBusy = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private void LoadGeppoFile()
        {
            var openFileDialog = new OpenFileDialog { Title = "編集する実績月報ファイルを選択", Filter = "Excel ファイル (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                var result = MessageBox.Show("選択したファイルの内容で現在の作業内容を上書きします。\nよろしいですか？（現在の入力内容は失われます）", "上書き確認", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Cancel) return;
                try
                {
                    File.Copy(openFileDialog.FileName, InputFilePath, true);
                    ReloadAllData();
                    Log($"実績月報 '{Path.GetFileName(openFileDialog.FileName)}' を読み込みました。");
                    MessageBox.Show("実績月報のデータを読み込みました。", "読み込み完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "実績月報ファイルの読み込み中にエラーが発生しました。");
                    MessageBox.Show($"ファイルの読み込みに失敗しました。\n詳細はログファイルを確認してください。", "読み込みエラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveInputFile()
        {
            try { _excelHandler.Save(); MessageBox.Show($"現在の入力内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information); Log($"--- 入力内容を保存しました ---"); }
            catch (Exception ex) { Logger.Error(ex, "入力内容の保存中にエラーが発生しました。"); MessageBox.Show($"保存に失敗しました。\n詳細はログファイルを確認してください。", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void ClearInputData(bool showSuccessMessage)
        {
            Log("--- 入力データをクリアします ---");
            var logMessages = _excelHandler.ClearData();
            foreach (var msg in logMessages) Log(msg);
            _registeredEastSheets.Clear();
            UpdateEastSheetStatus();
            _excelHandler.Save();
            UpdatePreview();
            if (showSuccessMessage) { MessageBox.Show("入力データをクリアしました。", "クリア完了", MessageBoxButton.OK, MessageBoxImage.Information); }
        }

        private async Task CheckForUpdate()
        {
            var updateService = new UpdateService(CurrentVersion, GithubToken, VersionInfoUrl, ReleasesPageUrl, Log);
            await updateService.CheckForUpdateAsync();
        }

        private void Log(string message)
        {
            Logger.Info(message);
            void updateAction() { _logBuilder.AppendLine(message); LogText = _logBuilder.ToString(); }
            if (Application.Current.Dispatcher.CheckAccess()) { updateAction(); }
            else { Application.Current.Dispatcher.Invoke(updateAction); }
        }

        private static bool TryParseValue(string input, string fieldName, out double? result, bool silent = false)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input)) return true;
            if (double.TryParse(input, out double parsedValue)) { result = parsedValue; return true; }
            if (!silent)
            {
                MessageBox.Show($"「{input}」は {fieldName} の数値として認識できません。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return false;
        }

        private void OpenSettings()
        {
            var settingsVM = new SettingsWindowViewModel(Rates, _excelHandler, RatesFilePath, this, _shortcutService);
            var settingsWindow = new SettingsWindow(settingsVM) { Owner = Application.Current.MainWindow };
            settingsWindow.ShowDialog();
        }

        private void OpenHelp()
        {
            try
            {
                if (File.Exists(HelpFilePath)) { Process.Start(new ProcessStartInfo(HelpFilePath) { UseShellExecute = true }); Log("ヘルプファイルを開きました。"); }
                else { MessageBox.Show("ヘルプファイル (readme.pdf) が見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Logger.Warn("ヘルプファイルが見つかりませんでした: " + HelpFilePath); }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ヘルプファイルを開けませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Logger.Error(ex, "ヘルプファイルのオープン中にエラーが発生しました。");
            }
        }

        private void OpenEditWindow()
        {
            if (SelectedRow == null) return;
            var editVM = new EditWindowViewModel(this, SelectedNormalSheet, SelectedRow);
            var editWindow = new EditWindow(editVM) { Owner = Application.Current.MainWindow };
            editWindow.ShowDialog();
        }

        private void DeleteSelectedRow()
        {
            if (SelectedRow == null) return;

            var sheet = SelectedNormalSheet;
            var rowIndex = SelectedRow.RowIndex;

            var result = MessageBox.Show($"選択した行({rowIndex}行目)を削除しますか？\nこの操作は元に戻せません。", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                _excelHandler.DeleteRows(sheet, new List<int> { rowIndex });
                UpdatePreview();
                _excelHandler.Save();
                Log($"[{sheet}] から {rowIndex}行目のデータを削除しました。");
            }
        }
    }
}