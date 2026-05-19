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
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.ViewModels
{
    public class MainViewModel : ObservableObject, IDisposable
    {
        private bool _disposed = false;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        #region 定数・パス

        private const string VersionInfoUrl = "https://raw.githubusercontent.com/ligdoor/HansoInputTool/refs/heads/master/HansoInputTool/version.json";
        private const string ReleasesPageUrl = "https://github.com/ligdoor/HansoInputTool/releases";
        private const int MaxLogLines = 200;

        // データパスは App.OnStartup で DataSetupService によって確定済み
        private static string BaseDataPath               => App.DataPath
                                                            ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        private static readonly string VersionFilePath   = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "version.json");
        private static string RatesFilePath              => Path.Combine(BaseDataPath, "rates.json");
        private static string InputFilePath              => Path.Combine(BaseDataPath, "Input.xlsx");
        private static string TemplateFilePath           => Path.Combine(BaseDataPath, "Template.xlsx");
        private static string ColumnMapFilePath          => Path.Combine(BaseDataPath, "column_map.json");
        private static string CustomFlagsFilePath        => Path.Combine(BaseDataPath, "custom_flags.json");
        private static string DatabaseFilePath           => Path.Combine(BaseDataPath, "hanso_data.db");
        private static string HelpFilePath               => Path.Combine(BaseDataPath, "readme.pdf");
        private static string ShortcutSettingsFilePath   => Path.Combine(BaseDataPath, "shortcuts.json");

        #endregion

        #region サービス

        private readonly BackupService _backupService;
        private readonly ValidationService _validationService;
        private ShortcutService _shortcutService;
        private ExcelHandler _excelHandler;
        private DatabaseService _dbService;
        private ColumnMapping _columnMap;
        private List<string> _allSheetNames;
        private FlagDefinitionService _flagService;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ShortcutService ShortcutService => _shortcutService;
        public FlagDefinitionService FlagService => _flagService;

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
        public ICollectionView PreviewDataView { get; }

        private RowData _selectedRow;
        public RowData SelectedRow { get => _selectedRow; set => SetProperty(ref _selectedRow, value); }

        private string _period;
        public string Period { get => _period; set => SetProperty(ref _period, value); }

        private string _month;
        public string Month { get => _month; set => SetProperty(ref _month, value); }

        private string _rNumber;
        public string RNumber { get => _rNumber; set => SetProperty(ref _rNumber, value); }

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

        // XAMLバインディング互換のため子VMのコマンドを公開
        public ICommand RegisterNormalCommand => NormalSheet.RegisterCommand;
        public ICommand RegisterEastCommand   => EastSheet.RegisterCommand;

        #endregion

        public MainViewModel()
        {
            _backupService     = new BackupService();
            _validationService = new ValidationService();
            NormalSheet = new NormalSheetViewModel(_validationService);
            EastSheet   = new EastSheetViewModel(_validationService);

            OpenSettingsCommand               = new RelayCommand(p => OpenSettings(),              p => !IsBusy);
            OpenHelpCommand                   = new RelayCommand(p => OpenHelp(),                  p => !IsBusy);
            CreateBackupCommand               = new RelayCommand(p => CreateManualBackup(),        p => !IsBusy);
            RestoreBackupCommand              = new RelayCommand(p => OpenRestoreBackupWindow(),   p => !IsBusy);
            OpenBackupFolderCommand           = new RelayCommand(p => _backupService.OpenBackupFolder(), p => !IsBusy);
            EditRowCommand                    = new RelayCommand(p => OpenEditWindow(),            p => SelectedRow != null && !IsBusy);
            DeleteRowCommand                  = new RelayCommand(p => DeleteSelectedRow(),         p => SelectedRow != null && !IsBusy);
            LoadGeppoFileCommand              = new RelayCommand(p => LoadGeppoFile(),             p => !IsBusy);
            SaveInputCommand                  = new RelayCommand(p => SaveInputFile(),             p => !IsBusy);
            TransferCommand                   = new RelayCommand(async p => await StartTransfer(), p => !IsBusy);
            OnLoadedCommand                   = new RelayCommand(async p => await OnWindowLoaded());
            OnClosingCommand                  = new RelayCommand(p => { });
            OpenMonthlyReportDashboardCommand = new RelayCommand(_ => OpenWindow<MonthlyReportDashboardWindow>("月報統計ダッシュボード"));
            OpenVehicleAnnualSummaryCommand   = new RelayCommand(_ => OpenWindow<VehicleAnnualSummaryWindow>("車両別年度集計"));
            OpenPdfImportCommand              = new RelayCommand(_ => OpenPdfImport(), _ => !IsBusy);

            PreviewDataView = CollectionViewSource.GetDefaultView(PreviewData);
        }

        #region 初期化

        private async Task OnWindowLoaded()
        {
            try
            {
                Logger.Info("アプリケーションの初期化を開始します。");

                if (!Directory.Exists(BaseDataPath))
                {
                    MessageBox.Show("データフォルダが見つかりません。\n実行ファイルと同じ場所に 'data' フォルダを配置してください。",
                        "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                    return;
                }

                _backupService.CreateAutoBackup(InputFilePath);
                _backupService.CreateAutoBackup(TemplateFilePath);
                Log("起動時の自動バックアップを作成しました。");

                var ratesJson = await File.ReadAllTextAsync(RatesFilePath);
                Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(ratesJson);

                var columnMapJson = await File.ReadAllTextAsync(ColumnMapFilePath);
                _columnMap = JsonConvert.DeserializeObject<ColumnMapping>(columnMapJson);

                _flagService     = new FlagDefinitionService(CustomFlagsFilePath);
                _excelHandler    = new ExcelHandler(InputFilePath, TemplateFilePath, _columnMap);
                _shortcutService = new ShortcutService(ShortcutSettingsFilePath);

                // SQLiteサービスを初期化してExcelHandlerに注入
                _dbService = new DatabaseService(DatabaseFilePath);
                _excelHandler.DbService = _dbService;

                // 起動時フラグ自動同期
                _excelHandler.SyncFlagsOnStartup(_flagService);
                Log("ショートカット設定を読み込みました。");

                // 月末日チェック用に年・月を渡す（Month は "1"〜"12" の文字列）
                NormalSheet.Initialize(_excelHandler, Log, UpdatePreview, _flagService,
                    getYearMonth: () =>
                    {
                        int.TryParse(Month, out var m);
                        return (DateTime.Now.Year, m);
                    });
                _excelHandler.FlagService = _flagService;
                EastSheet.Initialize(_excelHandler, Log);

                ReloadAllData();
                await CheckForUpdate();

                if (_excelHandler.CheckRemainingData())
                {
                    var result = MessageBox.Show("前回のデータが残っています。\n全ての入力データをクリアして新規に開始しますか？",
                        "データクリア確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                        ClearInputData(true);
                }

                Logger.Info("アプリケーションの初期化が完了しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "アプリケーションの初期化中に致命的なエラーが発生しました。");
                MessageBox.Show($"初期化エラー:\n{ex.GetType().Name}\n{ex.Message}\n\n内部エラー:{ex.InnerException?.Message}",
                    "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        #endregion

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

        #endregion

        #region ショートカット

        public bool ProcessShortcut(Key key, ModifierKeys modifiers)
        {
            if (_shortcutService == null) return false;
            foreach (var kvp in _shortcutService.CurrentSettings.Shortcuts)
                if (kvp.Value.Matches(key, modifiers))
                    return ExecuteShortcutAction(kvp.Key);
            return false;
        }

        private bool ExecuteShortcutAction(string actionName)
        {
            if (IsBusy) return false;
            switch (actionName)
            {
                case "Save":         return TryExecute(SaveInputCommand);
                case "Register":     return SelectedTabIndex == 0 ? TryExecute(RegisterNormalCommand) : TryExecute(RegisterEastCommand);
                case "NextSheet":    MoveSheet(+1); return true;
                case "PrevSheet":    MoveSheet(-1); return true;
                case "Transfer":     return TryExecute(TransferCommand);
                case "OpenSettings": return TryExecute(OpenSettingsCommand);
                case "SwitchTab":    SelectedTabIndex = (SelectedTabIndex + 1) % 2; return true;
                case "EditRow":      return TryExecute(EditRowCommand);
                case "DeleteRow":    return TryExecute(DeleteRowCommand);
                case "CreateBackup": return TryExecute(CreateBackupCommand);
            }
            return false;
        }

        private static bool TryExecute(ICommand cmd)
        {
            if (cmd.CanExecute(null)) { cmd.Execute(null); return true; }
            return false;
        }

        private void MoveSheet(int direction)
        {
            if (SelectedTabIndex == 0)
                MoveInCollection(NormalSheet.NormalSheets, NormalSheet.SelectedNormalSheet, s => NormalSheet.SelectedNormalSheet = s, direction);
            else
                MoveInCollection(EastSheet.EastSheets, EastSheet.SelectedEastSheet, s => EastSheet.SelectedEastSheet = s, direction);
        }

        private static void MoveInCollection(ObservableCollection<string> list, string current, Action<string> setter, int direction)
        {
            if (list.Count == 0) return;
            int index = list.IndexOf(current) + direction;
            if (index < 0) index = list.Count - 1;
            else if (index >= list.Count) index = 0;
            setter(list[index]);
        }

        #endregion

        #region 転記

        private async Task StartTransfer()
        {
            if (!int.TryParse(Period, out var period) || !int.TryParse(Month, out var month) || !int.TryParse(RNumber, out var rNum))
            {
                MessageBox.Show("期、月、R番号を正しく入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool shouldContinue = false;
            var confirmVM = new TransferConfirmationViewModel(_excelHandler, Rates, _columnMap, Period, Month, RNumber, r => shouldContinue = r, _flagService);
            new TransferConfirmationWindow(confirmVM) { Owner = Application.Current.MainWindow }.ShowDialog();
            if (!shouldContinue) { Log("転記処理がキャンセルされました。"); return; }

            var dialog = new OpenFileDialog
            {
                Title = "出力先のベースフォルダを選択してください",
                CheckFileExists = false, CheckPathExists = true,
                FileName = "フォルダを選択", Filter = "Folder|.", ValidateNames = false, DereferenceLinks = true
            };
            if (dialog.ShowDialog() != true) { Log("フォルダ選択がキャンセルされました。"); return; }
            string outputDir = Path.GetDirectoryName(dialog.FileName);

            IsBusy = true;
            var progressVM = new ProgressWindowViewModel();
            var progressWindow = new ProgressWindow(progressVM) { Owner = Application.Current.MainWindow };
            var progress = new Progress<TransferProgressReport>(report =>
            {
                if (report.Total > 0)
                    progressVM.UpdateProgress(report.Current, report.Total, report.Message ?? "");
                else if (!string.IsNullOrEmpty(report.Message))
                    progressVM.AppendLog(report.Message);
            });
            progressWindow.Show();

            try
            {
                _excelHandler.Save();
                await new TransferService().ExecuteAsync(
                    InputFilePath, TemplateFilePath, outputDir,
                    period, month, rNum, _allSheetNames, Rates, _columnMap, progress, _flagService, _dbService);

                Log("========\n転記完了\n========");
                Period = Month = RNumber = string.Empty;
                progressVM.Complete("2つのファイルの作成が完了しました。");
                if (_dbService != null)
                {
                    _excelHandler.InvalidateCacheAll();
                    EastSheet.ClearRegisteredSheets();
                    UpdatePreview();
                }
                else
                {
                    ClearInputData(false);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "転記処理中にエラーが発生しました。");
                Log($"エラー: {ex.Message}");
                progressVM.ErrorComplete("エラーが発生しました: 詳細はログファイルを確認してください。");
            }
            finally
            {
                IsBusy = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }

        #endregion

        #region ファイル操作

        private void LoadGeppoFile()
        {
            var dialog = new OpenFileDialog { Title = "編集する実績月報ファイルを選択", Filter = "Excel ファイル (*.xlsx)|*.xlsx" };
            if (dialog.ShowDialog() != true) return;
            if (MessageBox.Show("選択したファイルの内容で現在の作業内容を上書きします。\nよろしいですか？",
                    "上書き確認", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.Cancel) return;
            try
            {
                File.Copy(dialog.FileName, InputFilePath, true);
                ReloadAllData();
                Log($"実績月報 '{Path.GetFileName(dialog.FileName)}' を読み込みました。");
                MessageBox.Show("実績月報のデータを読み込みました。", "読み込み完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "実績月報ファイルの読み込み中にエラーが発生しました。");
                MessageBox.Show("ファイルの読み込みに失敗しました。\n詳細はログファイルを確認してください。",
                    "読み込みエラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveInputFile()
        {
            try
            {
                _excelHandler.Save();
                MessageBox.Show("現在の入力内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                Log("--- 入力内容を保存しました ---");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "入力内容の保存中にエラーが発生しました。");
                MessageBox.Show("保存に失敗しました。\n詳細はログファイルを確認してください。",
                    "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region バックアップ

        private void CreateManualBackup()
        {
            try
            {
                var inputBackup    = _backupService.CreateManualBackup(InputFilePath,    "手動保存");
                var templateBackup = _backupService.CreateManualBackup(TemplateFilePath, "手動保存");
                if (inputBackup != null && templateBackup != null)
                {
                    MessageBox.Show(
                        $"バックアップを作成しました。\n\nInput.xlsx: {Path.GetFileName(inputBackup)}\nTemplate.xlsx: {Path.GetFileName(templateBackup)}\n\n保存場所: backupsフォルダ",
                        "バックアップ完了", MessageBoxButton.OK, MessageBoxImage.Information);
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
                MessageBox.Show("バックアップの作成に失敗しました。\n詳細はログファイルを確認してください。",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenRestoreBackupWindow()
        {
            var vm = new RestoreBackupWindowViewModel(_backupService, InputFilePath, TemplateFilePath, this);
            new RestoreBackupWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        #endregion

        #region ウィンドウ管理

        private void OpenWindow<T>(string displayName) where T : Window, new()
        {
            try
            {
                Logger.Info($"{displayName}ウィンドウを開きます");
                new T { Owner = Application.Current.MainWindow }.Show();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"{displayName}ウィンドウを開く際にエラーが発生");
                MessageBox.Show($"ウィンドウを開けませんでした: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenSettings()
        {
            var vm = new SettingsWindowViewModel(Rates, _excelHandler, RatesFilePath, this, _shortcutService, _backupService, _flagService);
            new SettingsWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private void OpenPdfImport()
        {
            var apiKey = LoadApiKey();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                var inputDialog = new ApiKeyInputWindow { Owner = Application.Current.MainWindow };
                if (inputDialog.ShowDialog() != true) return;
                apiKey = inputDialog.ApiKey;
                if (!string.IsNullOrWhiteSpace(apiKey))
                    SaveApiKey(apiKey);
            }

            var vm = new PdfImportViewModel(NormalSheet, Log, apiKey);
            new Views.PdfImportWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private static readonly string ApiSettingsFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "data", "api_settings.json");

        // AES暗号化用の固定キー（変更しないこと）
        private static readonly byte[] AesKey = System.Text.Encoding.UTF8.GetBytes("HansoTool!AES256Key#2025$Secure!"); // 32バイト
        private static readonly byte[] AesIv  = System.Text.Encoding.UTF8.GetBytes("HansoIV!16Bytes!"); // 16バイト

        private string LoadApiKey()
        {
            try
            {
                if (!File.Exists(ApiSettingsFilePath)) return null;
                var json = File.ReadAllText(ApiSettingsFilePath);
                var obj = JObject.Parse(json);
                var encrypted = obj["claude_api_key"]?.ToString();
                if (string.IsNullOrEmpty(encrypted)) return null;

                using var aes = System.Security.Cryptography.Aes.Create();
                aes.Key = AesKey;
                aes.IV  = AesIv;
                using var decryptor = aes.CreateDecryptor();
                var encryptedBytes = Convert.FromBase64String(encrypted);
                using var ms = new MemoryStream(encryptedBytes);
                using var cs = new System.Security.Cryptography.CryptoStream(ms, decryptor, System.Security.Cryptography.CryptoStreamMode.Read);
                using var reader = new StreamReader(cs);
                return reader.ReadToEnd();
            }
            catch { return null; }
        }

        private void SaveApiKey(string apiKey)
        {
            try
            {
                using var aes = System.Security.Cryptography.Aes.Create();
                aes.Key = AesKey;
                aes.IV  = AesIv;
                using var encryptor = aes.CreateEncryptor();
                using var ms = new MemoryStream();
                using var cs = new System.Security.Cryptography.CryptoStream(ms, encryptor, System.Security.Cryptography.CryptoStreamMode.Write);
                using var writer = new StreamWriter(cs);
                writer.Write(apiKey);
                writer.Flush();
                cs.FlushFinalBlock();
                var encrypted = Convert.ToBase64String(ms.ToArray());

                var obj = new JObject { ["claude_api_key"] = encrypted };
                File.WriteAllText(ApiSettingsFilePath, obj.ToString());
            }
            catch (Exception ex) { Logger.Warn(ex, "APIキーの保存に失敗しました"); }
        }

        private void OpenHelp()
        {
            try
            {
                if (File.Exists(HelpFilePath))
                    { Process.Start(new ProcessStartInfo(HelpFilePath) { UseShellExecute = true }); Log("ヘルプファイルを開きました。"); }
                else
                    MessageBox.Show("ヘルプファイル (readme.pdf) が見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
            var vm = new EditWindowViewModel(this, NormalSheet.SelectedNormalSheet, SelectedRow);
            new EditWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private void DeleteSelectedRow()
        {
            if (SelectedRow == null) return;
            var sheet    = NormalSheet.SelectedNormalSheet;
            var rowIndex = SelectedRow.RowIndex;
            int idToDelete = (SelectedRow.DbId > 0) ? (int)SelectedRow.DbId : rowIndex;
            if (MessageBox.Show($"選択した行({rowIndex}行目)を削除しますか？\nこの操作は元に戻せません。",
                    "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                _excelHandler.DeleteRows(sheet, new List<int> { idToDelete });
                UpdatePreview();
                if (_dbService == null) _excelHandler.Save();
                Log($"[{sheet}] から {rowIndex}行目のデータを削除しました。");
            }
        }

        #endregion

        #region バージョン確認

        private async Task CheckForUpdate()
        {
            string currentVersion = "0.0.0";
            try
            {
                if (File.Exists(VersionFilePath))
                {
                    var versionData = JObject.Parse(await File.ReadAllTextAsync(VersionFilePath));
                    currentVersion = versionData["latest_version"]?.ToString() ?? "0.0.0";
                    AppVersion = $"v{currentVersion}";
                    Logger.Info($"ローカルバージョン: {currentVersion}");
                }
                else Logger.Warn($"version.json が見つかりません: {VersionFilePath}");
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "version.json の読み込みに失敗しました。バージョンチェックをスキップします。");
                return;
            }
            await new UpdateService(currentVersion, "", VersionInfoUrl, ReleasesPageUrl, Log).CheckForUpdateAsync();
        }

        #endregion

        #region ログ

        private void Log(string message)
        {
            Logger.Info(message);
            void Update()
            {
                _logBuilder.AppendLine(message);
                var lines = _logBuilder.ToString().Split('\n');
                if (lines.Length > MaxLogLines)
                {
                    _logBuilder.Clear();
                    _logBuilder.Append(string.Join("\n", lines.Skip(lines.Length - MaxLogLines)));
                }
                LogText = _logBuilder.ToString();
            }
            if (Application.Current.Dispatcher.CheckAccess()) Update();
            else Application.Current.Dispatcher.Invoke(Update);
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed) { _excelHandler?.Dispose(); _dbService?.Dispose(); _disposed = true; }
        }
    }
}
