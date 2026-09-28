using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Models;
using Microsoft.Win32;

namespace HansoInputTool.ViewModels
{
    public class PdfImportViewModel : ObservableObject
    {
        private readonly NormalSheetViewModel _normalSheet;
        private readonly Action<string> _log;
        private readonly AiSettings _aiSettings;

        public ObservableCollection<PdfImportItem> Items { get; } = new();
        public ObservableCollection<string> VehicleSheets { get; } = new();

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set { if (SetProperty(ref _isBusy, value)) CommandManager.InvalidateRequerySuggested(); }
        }

        private string _statusMessage = "PDFを選択してください";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private int _progressCurrent;
        public int ProgressCurrent { get => _progressCurrent; set => SetProperty(ref _progressCurrent, value); }

        private int _progressTotal;
        public int ProgressTotal { get => _progressTotal; set => SetProperty(ref _progressTotal, value); }

        public bool HasItems => Items.Count > 0;

        public ICommand SelectAndAnalyzePdfCommand { get; }
        public ICommand RegisterAllCommand         { get; }
        public ICommand RegisterItemCommand        { get; }
        public ICommand RemoveItemCommand          { get; }
        public ICommand OpenPageCommand            { get; }
        public string ProviderName => _aiSettings?.Provider ?? "未設定";
        public string ModelName => _aiSettings?.Model ?? "";

        public PdfImportViewModel(NormalSheetViewModel normalSheet, Action<string> log, AiSettings aiSettings)
        {
            _normalSheet = normalSheet;
            _log         = log;
            _aiSettings  = aiSettings ?? new AiSettings();
            foreach (var sheet in _normalSheet.NormalSheets) VehicleSheets.Add(sheet);

            SelectAndAnalyzePdfCommand = new RelayCommand(async _ => await SelectAndAnalyzeAsync(), _ => !IsBusy);
            RegisterAllCommand         = new RelayCommand(async _ => await RegisterAllAsync(),       _ => !IsBusy && Items.Any(i => i.CanRegister));
            RegisterItemCommand        = new RelayCommand(async p => await RegisterItemAsync(p as PdfImportItem), p => !IsBusy && (p as PdfImportItem)?.CanRegister == true);
            RemoveItemCommand          = new RelayCommand(p => RemoveItem(p as PdfImportItem), p => p != null && !IsBusy);
            OpenPageCommand            = new RelayCommand(p => OpenPage(p as PdfImportItem), p => p is PdfImportItem && !IsBusy);
        }

        private async Task SelectAndAnalyzeAsync()
        {
            var dialog = new OpenFileDialog
            {
                Title  = "日報PDFを選択",
                Filter = "PDFファイル (*.pdf)|*.pdf",
                Multiselect = false
            };
            if (dialog.ShowDialog() != true) return;

            IsBusy = true;
            Items.Clear();
            OnPropertyChanged(nameof(HasItems));

            try
            {
                using var ocrService = new PdfOcrService();
                var pages = await ocrService.AnalyzeAllPagesAsync(
                    dialog.FileName,
                    _aiSettings,
                    (current, total) =>
                    {
                        ProgressCurrent = current;
                        ProgressTotal   = total;
                        StatusMessage   = $"解析中... {current}/{total}ページ";
                    });

                foreach (var data in pages)
                {
                    var item = new PdfImportItem(data, _normalSheet.NormalSheets.ToList());
                    item.PropertyChanged += (_, e) =>
                    {
                        if (e.PropertyName == nameof(PdfImportItem.HasError)
                            || e.PropertyName == nameof(PdfImportItem.IsConfirmed)
                            || e.PropertyName == nameof(PdfImportItem.IsDone))
                            RefreshResultSummary();
                    };
                    Items.Add(item);
                }

                OnPropertyChanged(nameof(HasItems));
                RefreshResultSummary();

                _log?.Invoke($"[PDF読込] {Path.GetFileName(dialog.FileName)}: {pages.Count}ページ解析完了");
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
                _log?.Invoke($"[PDF読込エラー] {ex.Message}");
                MessageBox.Show($"PDF解析中にエラーが発生しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task RegisterAllAsync()
        {
            var targets = Items.Where(i => i.CanRegister).ToList();
            IsBusy = true;
            int success = 0;

            foreach (var item in targets)
            {
                if (await DoRegisterAsync(item)) success++;
            }

            IsBusy = false;
            StatusMessage = $"登録完了: {success}/{targets.Count}件";
            _log?.Invoke($"[PDF一括登録] {success}件登録しました。");
        }

        private async Task RegisterItemAsync(PdfImportItem item)
        {
            if (item == null) return;
            IsBusy = true;
            await DoRegisterAsync(item);
            IsBusy = false;
        }

        private async Task<bool> DoRegisterAsync(PdfImportItem item)
        {
            try
            {
                item.StatusText = "⏳ 登録中...";

                _normalSheet.Day       = item.Day;
                _normalSheet.YuryoKm  = item.YuryoKm;
                _normalSheet.MuryoKm  = item.MuryoKm;
                _normalSheet.LateValue = item.ShinyaMinutes;
                _normalSheet.SelectedNormalSheet = item.MatchedVehicle;
                _normalSheet.HansoCountOverride = item.WorkType == "搬送" ? 1 : 0;
                _normalSheet.IsFuelChecked = item.FuelMarked;
                _normalSheet.FuelLiters = item.FuelLiters;
                _normalSheet.FuelOdometerKm = item.FuelOdometerKm;
                _normalSheet.ResetFlags();
                if (item.EmbalingConfirmed)
                {
                    var embalmingFlag = _normalSheet.FlagItems.FirstOrDefault(f => f.Id == "embalming" || f.DisplayName.Contains("エンバー") || f.DisplayName.Contains("エンバーミング"));
                    if (embalmingFlag != null) embalmingFlag.IsChecked = true;
                }

                if (await _normalSheet.RegisterPdfImportAsync())
                {
                    item.IsDone = true;
                    item.StatusText = "✅ 登録済み";
                    return true;
                }
                item.StatusText = "❌ 登録されませんでした。入力内容や確認メッセージを確認してください";
                return false;
            }
            catch (Exception ex)
            {
                item.StatusText = $"❌ 失敗: {ex.Message}";
                _log?.Invoke($"[登録エラー] {item.Label}: {ex.Message}");
                return false;
            }
        }

        private void RefreshResultSummary()
        {
            int needsFix = Items.Count(i => i.HasError);
            int unconfirmed = Items.Count(i => !i.IsConfirmed);
            int registered = Items.Count(i => i.IsDone);
            int readFailures = Items.Count(i => i.Data.RetryFailed && i.HasCoreError);
            StatusMessage = $"解析: {Items.Count}件 / AI読取失敗: {readFailures}件 / 要修正: {needsFix}件 / 未確認: {unconfirmed}件 / 登録済: {registered}件";
        }

        private void RemoveItem(PdfImportItem item)
        {
            if (item == null) return;
            Items.Remove(item);
            OnPropertyChanged(nameof(HasItems));
            RefreshResultSummary();
        }

        private void OpenPage(PdfImportItem item)
        {
            if (item?.Data?.PagePdfBytes == null || item.Data.PagePdfBytes.Length == 0) return;
            try
            {
                var dir = Path.Combine(Path.GetTempPath(), "HansoInputTool", "PdfReview");
                Directory.CreateDirectory(dir);
                var path = Path.Combine(dir, $"日報_p{item.Data.PageNumber}_{Guid.NewGuid():N}.pdf");
                File.WriteAllBytes(path, item.Data.PagePdfBytes);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                _log?.Invoke($"[PDF確認] ページ{item.Data.PageNumber}を開きました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDFを開けませんでした。\n\n{ex.Message}", "PDF確認", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    public class PdfImportItem : ObservableObject
    {
        public NippoData Data { get; }
        public string Label => $"p.{Data.PageNumber}  {(Data.Day.HasValue ? $"{Data.Day}日" : "日付不明")}  車両{Data.VehicleNumber ?? "?"}";

        private string _day;
        public string Day { get => _day; set { if (SetProperty(ref _day, value)) RaiseValidation(); } }
        private string _yuryoKm;
        public string YuryoKm { get => _yuryoKm; set { if (SetProperty(ref _yuryoKm, value)) RaiseValidation(); } }
        private string _muryoKm;
        public string MuryoKm { get => _muryoKm; set { if (SetProperty(ref _muryoKm, value)) RaiseValidation(); } }
        private string _shinyaMinutes;
        public string ShinyaMinutes { get => _shinyaMinutes; set { if (SetProperty(ref _shinyaMinutes, value)) RaiseValidation(); } }
        private string _vehicleNumber;
        public string VehicleNumber { get => _vehicleNumber; set => SetProperty(ref _vehicleNumber, value); }
        private string _matchedVehicle;
        public string MatchedVehicle { get => _matchedVehicle; set { if (SetProperty(ref _matchedVehicle, value)) { OnPropertyChanged(nameof(HasVehicleMatch)); OnPropertyChanged(nameof(VehicleMatchText)); RaiseValidation(); } } }
        public bool HasVehicleMatch => !string.IsNullOrEmpty(MatchedVehicle);
        public string VehicleMatchText => HasVehicleMatch ? $"登録先: {MatchedVehicle}" : "登録先が一意に特定できません";
        private string _workType;
        public string WorkType { get => _workType; set { if (SetProperty(ref _workType, value)) RaiseValidation(); } }
        private bool _embalingConfirmed;
        public bool EmbalingConfirmed { get => _embalingConfirmed; set => SetProperty(ref _embalingConfirmed, value); }
        private bool _fuelMarked;
        public bool FuelMarked { get => _fuelMarked; set { if (SetProperty(ref _fuelMarked, value)) RaiseValidation(); } }
        private string _fuelMarkStatus;
        public string FuelMarkStatus
        {
            get => _fuelMarkStatus;
            set
            {
                if (SetProperty(ref _fuelMarkStatus, value))
                {
                    var marked = value == "給油あり";
                    if (_fuelMarked != marked)
                    {
                        _fuelMarked = marked;
                        OnPropertyChanged(nameof(FuelMarked));
                    }
                    RaiseValidation();
                }
            }
        }
        public string[] FuelMarkOptions { get; } = new[] { "未判定", "給油あり", "給油なし" };
        private string _fuelLiters;
        public string FuelLiters { get => _fuelLiters; set { if (SetProperty(ref _fuelLiters, value)) RaiseValidation(); } }
        private string _fuelOdometerKm;
        public string FuelOdometerKm { get => _fuelOdometerKm; set { if (SetProperty(ref _fuelOdometerKm, value)) RaiseValidation(); } }
        private string _statusText;
        public string StatusText
        {
            get
            {
                if (IsDone) return _statusText ?? "✅ 登録済み";
                if (_statusText?.StartsWith("⏳") == true || (_statusText?.StartsWith("❌") == true && !Data.RetryFailed)) return _statusText;
                var issues = GetValidationIssues();
                if (issues.Count > 0)
                {
                    string prefix = Data.RetryFailed ? $"❌ AI読取失敗（{RetryReason}） / 要修正: " : "⚠ 要修正: ";
                    return prefix + string.Join("、", issues);
                }
                if (!IsConfirmed) return "未確認: 左端の確認にチェック";
                return Data.RetryFailed ? "⚠ 読取失敗分を手入力で確認済み。登録できます" : "✅ 登録できます";
            }
            set
            {
                if (SetProperty(ref _statusText, value)) OnPropertyChanged(nameof(StatusText));
            }
        }
        private string RetryReason
        {
            get
            {
                var message = Data.RetryMessage ?? "";
                if (message.Contains("TooManyRequests", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("quota", StringComparison.OrdinalIgnoreCase)
                    || message.Contains("429", StringComparison.OrdinalIgnoreCase))
                    return "API利用上限。時間を置いて再解析";
                if (message.Contains("ServiceUnavailable", StringComparison.OrdinalIgnoreCase))
                    return "AI一時障害。時間を置いて再解析";
                return "応答を取得できません。項目を手入力して確認";
            }
        }
        public string EmbalmingCandidateText => Data.EmbalmingCandidate == true ? "候補検出: エンバー（チェックして確定）" : Data.EmbalmingCandidate == false ? "エンバー記載なし" : "エンバー判定不明（PDFで確認）";
        private bool _isDone;
        public bool IsDone { get => _isDone; set { if (SetProperty(ref _isDone, value)) { OnPropertyChanged(nameof(CanRegister)); OnPropertyChanged(nameof(HasError)); OnPropertyChanged(nameof(StatusText)); } } }
        private bool _isConfirmed;
        public bool IsConfirmed { get => _isConfirmed; set { if (SetProperty(ref _isConfirmed, value)) { OnPropertyChanged(nameof(CanRegister)); OnPropertyChanged(nameof(StatusText)); CommandManager.InvalidateRequerySuggested(); } } }

        public bool HasError => GetValidationIssues().Count > 0;
        public bool HasCoreError => !int.TryParse(Day, out var d) || d <= 0
            || !double.TryParse(YuryoKm, out _) || !double.TryParse(MuryoKm, out _);
        public bool CanRegister => !IsDone && IsConfirmed && !HasError;

        private List<string> GetValidationIssues()
        {
            var issues = new List<string>();
            if (!int.TryParse(Day, out var d) || d <= 0) issues.Add("日付");
            if (!double.TryParse(YuryoKm, out _)) issues.Add("有料km");
            if (!double.TryParse(MuryoKm, out _)) issues.Add("無料km");
            if (Data.RetryFailed && HasCoreError) issues.Insert(0, "コア項目を手入力");
            if (!int.TryParse(ShinyaMinutes, out var shinya) || shinya < 0) issues.Add("深夜分");
            if (WorkType != "搬送" && WorkType != "移動") issues.Add("搬送/移動");
            if (!HasVehicleMatch) issues.Add("登録車両");
            if (FuelMarkStatus != "給油あり" && FuelMarkStatus != "給油なし") issues.Add("給油の有無");
            if (FuelMarked && (!double.TryParse(FuelLiters, out var liters) || liters <= 0)) issues.Add("給油リッター");
            if (FuelMarked && (!double.TryParse(FuelOdometerKm, out var km) || km <= 0)) issues.Add("給油時距離");
            return issues;
        }

        public PdfImportItem(NippoData data, List<string> vehicleSheets)
        {
            Data = data;
            vehicleSheets ??= new List<string>();
            Day = data.Day?.ToString() ?? "";
            YuryoKm = data.YuryoKm?.ToString() ?? "";
            MuryoKm = data.MuryoKm?.ToString() ?? "";
            ShinyaMinutes = data.ShinyaMinutes?.ToString() ?? "";
            VehicleNumber = data.VehicleNumber ?? "";
            WorkType = data.WorkType ?? "";
            EmbalingConfirmed = false;
            FuelMarked = data.FuelMarked == true;
            FuelMarkStatus = data.FuelMarked == true ? "給油あり" : data.FuelMarked == false ? "給油なし" : "未判定";
            FuelLiters = data.FuelLiters?.ToString() ?? "";
            FuelOdometerKm = data.FuelOdometerKm?.ToString() ?? "";
            var digits = new string((data.VehicleNumber ?? "").Where(char.IsDigit).ToArray());
            var matches = string.IsNullOrWhiteSpace(digits) ? new List<string>() : vehicleSheets.Where(s => new string(s.Where(char.IsDigit).ToArray()).EndsWith(digits, StringComparison.Ordinal)).ToList();
            MatchedVehicle = matches.Count == 1 ? matches[0] : null;
            StatusText = data.RetryFailed ? $"❌ 読み取り失敗: {data.RetryMessage}" : null;
        }

        private void RaiseValidation()
        {
            if (_statusText?.StartsWith("❌") == true && !Data.RetryFailed) _statusText = null;
            OnPropertyChanged(nameof(HasError));
            OnPropertyChanged(nameof(HasCoreError));
            OnPropertyChanged(nameof(CanRegister));
            OnPropertyChanged(nameof(StatusText));
            CommandManager.InvalidateRequerySuggested();
        }
    }
}
