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
        private const string CurrentVersion = "0.2.5";
        private const string GithubToken = "";
        private const string VersionInfoUrl = "https://raw.githubusercontent.com/ligdoor/HansoInputToo/refs/heads/master/version.json";
        private const string ReleasesPageUrl = "https://github.com/ligdoor/HansoInputToo/releases";
        private static readonly string AppDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppName);
        private static readonly string RatesFilePath = Path.Combine(AppDataPath, "rates.json");
        private static readonly string WorkInputFilePath = Path.Combine(AppDataPath, "Input_work.xlsx");
        private static readonly string ColumnMapFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "column_map.json");
        private static readonly string BundledInputFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "Input.xlsx");
        private static readonly string BundledTemplateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "Template.xlsx");
        private static readonly string BundledRatesFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "rates.json");
        #endregion

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
        public string SelectedNormalSheet { get => _selectedNormalSheet; set { if (SetProperty(ref _selectedNormalSheet, value)) { UpdatePreview(); OnPropertyChanged(nameof(IsOotsukiSheet)); } } }
        public ObservableCollection<RowData> PreviewData { get; } = new();
        public ICollectionView PreviewDataView { get; }
        private RowData _selectedRow;
        public RowData SelectedRow { get => _selectedRow; set => SetProperty(ref _selectedRow, value); }
        public bool IsOotsukiSheet => SelectedNormalSheet?.Contains("大月") ?? false;
        private string _normalDay;
        public string NormalDay { get => _normalDay; set => SetProperty(ref _normalDay, value); }
        private string _normalYuryoKm;
        public string NormalYuryoKm { get => _normalYuryoKm; set => SetProperty(ref _normalYuryoKm, value); }
        private string _normalMuryoKm;
        public string NormalMuryoKm { get => _normalMuryoKm; set => SetProperty(ref _normalMuryoKm, value); }
        private string _normalLateValue;
        public string NormalLateValue { get => _normalLateValue; set => SetProperty(ref _normalLateValue, value); }
        private bool _isKoryo;
        public bool IsKoryo { get => _isKoryo; set => SetProperty(ref _isKoryo, value); }
        public ObservableCollection<string> EastSheets { get; } = new();
        private readonly List<string> _registeredEastSheets = new();
        private string _selectedEastSheet;
        public string SelectedEastSheet { get => _selectedEastSheet; set { if (SetProperty(ref _selectedEastSheet, value)) { UpdateEastSheetStatus(); } } }
        private string _eastSheetStatus = "（未登録）";
        public string EastSheetStatus { get => _eastSheetStatus; set => SetProperty(ref _eastSheetStatus, value); }
        private bool _isEastSheetRegistered = false;
        public bool IsEastSheetRegistered { get => _isEastSheetRegistered; set => SetProperty(ref _isEastSheetRegistered, value); }
        private string _eastJitsudo;
        public string EastJitsudo { get => _eastJitsudo; set => SetProperty(ref _eastJitsudo, value); }
        private string _eastHanso;
        public string EastHanso { get => _eastHanso; set => SetProperty(ref _eastHanso, value); }
        private string _eastYuryoKm;
        public string EastYuryoKm { get => _eastYuryoKm; set => SetProperty(ref _eastYuryoKm, value); }
        private string _eastMuryoKm;
        public string EastMuryoKm { get => _eastMuryoKm; set => SetProperty(ref _eastMuryoKm, value); }
        private string _eastUnso;
        public string EastUnso { get => _eastUnso; set => SetProperty(ref _eastUnso, value); }
        private string _period;
        public string Period { get => _period; set => SetProperty(ref _period, value); }
        private string _month;
        public string Month { get => _month; set => SetProperty(ref _month, value); }
        private string _rNumber;
        public string RNumber { get => _rNumber; set => SetProperty(ref _rNumber, value); }
        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }
        #endregion

        #region Commands
        public ICommand OpenSettingsCommand { get; }
        public ICommand RegisterNormalCommand { get; }
        public ICommand RegisterEastCommand { get; }
        public ICommand EditRowCommand { get; }
        public ICommand DeleteRowCommand { get; }
        public ICommand LoadGeppoFileCommand { get; }
        public ICommand SaveInputCommand { get; }
        public ICommand TransferCommand { get; }
        public ICommand OnLoadedCommand { get; }
        public ICommand OnClosingCommand { get; }
        #endregion

        public MainViewModel()
        {
            OpenSettingsCommand = new RelayCommand(OpenSettings, _ => !IsBusy);
            RegisterNormalCommand = new RelayCommand(RegisterNormal, _ => !IsBusy);
            RegisterEastCommand = new RelayCommand(RegisterEast, _ => !IsBusy);
            EditRowCommand = new RelayCommand(OpenEditWindow, _ => SelectedRow != null && !IsBusy);
            DeleteRowCommand = new RelayCommand(DeleteSelectedRow, _ => SelectedRow != null && !IsBusy);
            LoadGeppoFileCommand = new RelayCommand(_ => LoadGeppoFile(), _ => !IsBusy);
            SaveInputCommand = new RelayCommand(SaveInputFile, _ => !IsBusy);
            TransferCommand = new RelayCommand(async _ => await StartTransfer(), _ => !IsBusy);
            OnLoadedCommand = new RelayCommand(async _ => await OnWindowLoaded());
            OnClosingCommand = new RelayCommand(OnWindowClosing);

            PreviewDataView = CollectionViewSource.GetDefaultView(PreviewData);
        }

        private async Task OnWindowLoaded()
        {
            try
            {
                Logger.Info("アプリケーションの初期化を開始します。");
                Directory.CreateDirectory(AppDataPath);
                if (!File.Exists(RatesFilePath)) File.Copy(BundledRatesFilePath, RatesFilePath, true);
                if (!File.Exists(WorkInputFilePath)) File.Copy(BundledInputFilePath, WorkInputFilePath, true);
                var ratesJson = await File.ReadAllTextAsync(RatesFilePath);
                Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(ratesJson);
                var columnMapJson = await File.ReadAllTextAsync(ColumnMapFilePath);
                _columnMap = JsonConvert.DeserializeObject<ColumnMapping>(columnMapJson);
                _excelHandler = new ExcelHandler(WorkInputFilePath, _columnMap);
                _allSheetNames = _excelHandler.SheetNames;
                PopulateSheetCombos();
                await CheckForUpdate();
                if (_excelHandler.CheckRemainingData())
                {
                    var result = MessageBox.Show("前回のデータが残っています。\n全ての入力データをクリアして新規に開始しますか？", "データクリア確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes) { ClearInputData(true); }
                }
                UpdatePreview();
                Logger.Info("アプリケーションの初期化が完了しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "アプリケーションの初期化中に致命的なエラーが発生しました。");
                MessageBox.Show($"アプリケーションの初期化中にエラーが発生しました。\n詳細はログファイルを確認してください。", "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private void OnWindowClosing(object parameter)
        {
            if (parameter is not CancelEventArgs e) return;
            if (IsBusy) { MessageBox.Show("転記処理が実行中です。終了できません。", "処理中", MessageBoxButton.OK, MessageBoxImage.Warning); e.Cancel = true; return; }
            var result = MessageBox.Show("入力中のデータはファイルに保存されません。\nツールを終了しますか？", "終了確認", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (result == MessageBoxResult.Cancel) { e.Cancel = true; } else { if (File.Exists(WorkInputFilePath)) { try { File.Delete(WorkInputFilePath); } catch (Exception ex) { Logger.Warn(ex, "作業ファイルの削除に失敗しました。"); } } }
        }

        private void PopulateSheetCombos()
        {
            NormalSheets.Clear();
            EastSheets.Clear();
            _allSheetNames?.ForEach(s => { if (s.Contains("寝台車") || s.Contains("霊柩車")) NormalSheets.Add(s); if (s.Contains("東日本")) EastSheets.Add(s); });
            SelectedNormalSheet = NormalSheets.FirstOrDefault();
            SelectedEastSheet = EastSheets.FirstOrDefault();
        }

        private void UpdatePreview()
        {
            if (string.IsNullOrEmpty(SelectedNormalSheet)) { PreviewData.Clear(); return; }
            ;
            PreviewData.Clear();
            var data = _excelHandler.GetSheetDataForPreview(SelectedNormalSheet);
            foreach (var item in data) { PreviewData.Add(item); }
        }

        public void UpdateRowData(string sheetName, int rowIndex, Dictionary<string, double?> newValues, bool isKoryo)
        {
            _excelHandler.UpdateNormalData(sheetName, rowIndex, newValues, isKoryo);
            Log($"[{sheetName}] の {rowIndex}行目のデータを更新しました。（ファイル未保存）");
            UpdatePreview();
        }

        public void UpdateColumnMap(ColumnMapping newMap)
        {
            _columnMap = newMap;
            _excelHandler = new ExcelHandler(WorkInputFilePath, _columnMap);
            Log("列マッピング設定が更新されました。");
        }

        private async void RegisterNormal(object obj)
        {
            if (string.IsNullOrEmpty(SelectedNormalSheet)) { MessageBox.Show("通常シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (string.IsNullOrWhiteSpace(NormalDay)) { MessageBox.Show("日付は必須です。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var values = new Dictionary<string, double?>();
            if (!TryParseValue(NormalDay, "日(B)", out var dayVal)) return; values["日(B)"] = dayVal;
            if (!TryParseValue(NormalYuryoKm, "有料キロ(D)", out var yuryoKmVal)) return;
            values["有料キロ(D)"] = yuryoKmVal.HasValue ? Math.Round(yuryoKmVal.Value, MidpointRounding.AwayFromZero) : null;
            if (!TryParseValue(NormalMuryoKm, "無料キロ(E)", out var muryoKmVal)) return;
            values["無料キロ(E)"] = muryoKmVal.HasValue ? Math.Round(muryoKmVal.Value, MidpointRounding.AwayFromZero) : null;
            if (IsOotsukiSheet) { if (!TryParseValue(NormalLateValue, "深夜料金(H)", out var lateVal)) return; values["深夜料金(H)"] = lateVal; }
            else { if (!TryParseValue(NormalLateValue, "深夜時間(K)", out var lateVal)) return; values["深夜時間(K)"] = lateVal; }
            try
            {
                var (targetRow, insertInfo) = _excelHandler.RegisterNormalData(SelectedNormalSheet, values, IsKoryo);
                if (!string.IsNullOrEmpty(insertInfo)) Log($"[{SelectedNormalSheet}] {insertInfo}");
                Log($"[{SelectedNormalSheet}] の {targetRow}行目にデータを登録しました。(ファイル未保存)");
                NormalDay = NormalYuryoKm = NormalMuryoKm = NormalLateValue = string.Empty;
                IsKoryo = false;
                UpdatePreview();
                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "NormalDayTextBox" });
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "通常シートへのデータ登録中にエラーが発生しました。");
                MessageBox.Show($"登録エラーが発生しました。\n詳細はログファイルを確認してください。", "登録エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void RegisterEast(object obj)
        {
            if (string.IsNullOrEmpty(SelectedEastSheet)) { MessageBox.Show("東日本シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var values = new Dictionary<string, double?>();
            if (!TryParseValue(EastJitsudo, "延実働車輌数", out var jitsudo)) return; values["延実働車輌数"] = jitsudo;
            if (!TryParseValue(EastHanso, "搬送回数", out var hanso)) return; values["搬送回数"] = hanso;
            if (!TryParseValue(EastYuryoKm, "有料キロ数", out var yuryo)) return; values["有料キロ数"] = yuryo;
            if (!TryParseValue(EastMuryoKm, "無料キロ数", out var muryo)) return; values["無料キロ数"] = muryo;
            if (!TryParseValue(EastUnso, "運輸実績", out var unso)) return; values["運輸実績"] = unso;
            try
            {
                _excelHandler.RegisterEastData(SelectedEastSheet, values);
                Log($"[{SelectedEastSheet}] のデータを登録しました。(ファイル未保存)");
                if (!_registeredEastSheets.Contains(SelectedEastSheet)) { _registeredEastSheets.Add(SelectedEastSheet); }
                UpdateEastSheetStatus();
                EastJitsudo = EastHanso = EastYuryoKm = EastMuryoKm = EastUnso = string.Empty;
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
            if (!int.TryParse(Period, out var period) || !int.TryParse(Month, out var month) || !int.TryParse(RNumber, out var rNum)) { MessageBox.Show("期、月、R番号を正しく入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var dialog = new OpenFileDialog { Title = "出力先のベースフォルダを選択してください", CheckFileExists = false, CheckPathExists = true, FileName = "フォルダを選択", Filter = "Folder|.", ValidateNames = false, DereferenceLinks = true };
            if (dialog.ShowDialog() != true) { Log("フォルダ選択がキャンセルされました。"); return; }
            string outputDir = Path.GetDirectoryName(dialog.FileName);
            IsBusy = true;
            var progressVM = new ProgressWindowViewModel();
            var progressWindow = new ProgressWindow(progressVM) { Owner = Application.Current.MainWindow };
            var progress = new Progress<TransferProgressReport>(report => { if (!string.IsNullOrEmpty(report.Message)) { progressVM.AppendLog(report.Message); } if (report.Total > 0) { progressVM.UpdateProgress(report.Current, report.Total, ""); } });
            progressWindow.Show();
            try
            {
                _excelHandler.Save();
                var transferService = new TransferService();
                await transferService.ExecuteAsync(WorkInputFilePath, BundledTemplateFilePath, outputDir, period, month, rNum, _allSheetNames, Rates, _columnMap, progress);
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
            finally { IsBusy = false; CommandManager.InvalidateRequerySuggested(); }
        }

        private void LoadGeppoFile()
        {
            var openFileDialog = new OpenFileDialog { Title = "編集する実績月報ファイルを選択", Filter = "Excel ファイル (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                var result = MessageBox.Show("選択したファイルの内容で現在の作業内容を上書きします。\nよろしいですか？（保存していない入力内容は失われます）", "上書き確認", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Cancel) return;
                try
                {
                    File.Copy(openFileDialog.FileName, WorkInputFilePath, true);
                    _excelHandler.Load();
                    _allSheetNames = _excelHandler.SheetNames;
                    PopulateSheetCombos();
                    UpdatePreview();
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

        private void SaveInputFile(object obj)
        {
            try { _excelHandler.Save(); MessageBox.Show($"作業中の入力内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information); Log($"--- {WorkInputFilePath} を上書き保存しました ---"); }
            catch (Exception ex) { Logger.Error(ex, "入力内容の保存中にエラーが発生しました。"); MessageBox.Show($"保存に失敗しました。\n詳細はログファイルを確認してください。", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void ClearInputData(bool showSuccessMessage)
        {
            Log("--- Input.xlsx のデータクリア処理を開始 ---");
            var logMessages = _excelHandler.ClearData();
            foreach (var msg in logMessages) Log(msg);
            _registeredEastSheets.Clear();
            UpdateEastSheetStatus();
            SaveInputFile(null);
            _excelHandler.Load();
            UpdatePreview();
            if (showSuccessMessage) { MessageBox.Show("Input.xlsx の入力データをクリアし、保存しました。", "クリア完了", MessageBoxButton.OK, MessageBoxImage.Information); }
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

        private static bool TryParseValue(string input, string fieldName, out double? result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input)) return true;
            if (double.TryParse(input, out double parsedValue)) { result = parsedValue; return true; }
            MessageBox.Show($"「{input}」は {fieldName} の数値として認識できません。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private void OpenSettings(object obj)
        {
            var settingsVM = new SettingsWindowViewModel(Rates, _columnMap, RatesFilePath, ColumnMapFilePath, this);
            var settingsWindow = new SettingsWindow(settingsVM) { Owner = Application.Current.MainWindow };
            settingsWindow.ShowDialog();
        }

        private void OpenEditWindow(object obj)
        {
            if (SelectedRow == null) return;
            var editVM = new EditWindowViewModel(this, SelectedNormalSheet, SelectedRow);
            var editWindow = new EditWindow(editVM) { Owner = Application.Current.MainWindow };
            editWindow.ShowDialog();
        }

        private void DeleteSelectedRow(object obj)
        {
            if (SelectedRow == null) return;
            var result = MessageBox.Show($"選択した行({SelectedRow.RowIndex}行目)を削除しますか？\nこの操作は元に戻せません。", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                _excelHandler.DeleteRows(SelectedNormalSheet, new List<int> { SelectedRow.RowIndex });
                Log($"[{SelectedNormalSheet}] から {SelectedRow.RowIndex}行目のデータを削除しました。（ファイル未保存）");
                UpdatePreview();
            }
        }
    }
}