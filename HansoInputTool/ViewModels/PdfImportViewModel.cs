using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.Models;
using HansoInputTool.ViewModels.Base;
using Microsoft.Win32;

namespace HansoInputTool.ViewModels
{
    public class PdfImportViewModel : ObservableObject
    {
        private readonly NormalSheetViewModel _normalSheet;
        private readonly Action<string> _log;
        private readonly AiSettings _aiSettings;

        public ObservableCollection<PdfImportItem> Items { get; } = new();

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
                    var item = new PdfImportItem(data);
                    Items.Add(item);
                }

                OnPropertyChanged(nameof(HasItems));
                var errorCount = Items.Count(i => i.HasError);
                StatusMessage = errorCount > 0
                    ? $"解析完了: {Items.Count}件（うち{errorCount}件要確認）"
                    : $"解析完了: {Items.Count}件 — 内容を確認して登録してください";

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
                _normalSheet.LateValue = string.IsNullOrEmpty(item.ShinyaMinutes) ? "0" : item.ShinyaMinutes;
                _normalSheet.ResetFlags();

                await Task.Delay(100);

                if (_normalSheet.RegisterCommand.CanExecute(null))
                {
                    _normalSheet.RegisterCommand.Execute(null);
                    await Task.Delay(300);
                }

                item.IsDone     = true;
                item.StatusText = $"✅ 登録済み";
                return true;
            }
            catch (Exception ex)
            {
                item.StatusText = $"❌ 失敗: {ex.Message}";
                _log?.Invoke($"[登録エラー] {item.Label}: {ex.Message}");
                return false;
            }
        }

        private void RemoveItem(PdfImportItem item)
        {
            if (item == null) return;
            Items.Remove(item);
            OnPropertyChanged(nameof(HasItems));
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
        private string _statusText;
        public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }
        private bool _isDone;
        public bool IsDone { get => _isDone; set { if (SetProperty(ref _isDone, value)) { OnPropertyChanged(nameof(CanRegister)); OnPropertyChanged(nameof(HasError)); } } }
        private bool _isConfirmed;
        public bool IsConfirmed { get => _isConfirmed; set { if (SetProperty(ref _isConfirmed, value)) { OnPropertyChanged(nameof(CanRegister)); CommandManager.InvalidateRequerySuggested(); } } }

        public bool HasError => !int.TryParse(Day, out var d) || d <= 0
            || !double.TryParse(YuryoKm, out _) || !double.TryParse(MuryoKm, out _)
            || Data.RetryFailed;
        public bool CanRegister => !IsDone && IsConfirmed && !HasError;

        public PdfImportItem(NippoData data)
        {
            Data = data;
            Day = data.Day?.ToString() ?? "";
            YuryoKm = data.YuryoKm?.ToString() ?? "";
            MuryoKm = data.MuryoKm?.ToString() ?? "";
            ShinyaMinutes = data.ShinyaMinutes.HasValue ? data.ShinyaMinutes.ToString() : "0";
            VehicleNumber = data.VehicleNumber ?? "";
            StatusText = data.RetryFailed
                ? $"❌ 読み取り失敗: {data.RetryMessage}"
                : (data.ValidateRequired().isValid ? "⚠️ 未確認" : $"⚠️ 要確認: {data.ValidateRequired().missingFields}");
        }

        private void RaiseValidation()
        {
            OnPropertyChanged(nameof(HasError));
            OnPropertyChanged(nameof(CanRegister));
        }
    }
}
