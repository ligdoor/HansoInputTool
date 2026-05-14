using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using Microsoft.Win32;

namespace HansoInputTool.ViewModels
{
    public class PdfImportViewModel : ObservableObject
    {
        private readonly NormalSheetViewModel _normalSheet;
        private readonly Action<string> _log;
        private readonly string _apiKey;

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

        public PdfImportViewModel(NormalSheetViewModel normalSheet, Action<string> log, string apiKey)
        {
            _normalSheet = normalSheet;
            _log         = log;
            _apiKey      = apiKey;

            SelectAndAnalyzePdfCommand = new RelayCommand(async _ => await SelectAndAnalyzeAsync(), _ => !IsBusy);
            RegisterAllCommand         = new RelayCommand(async _ => await RegisterAllAsync(),       _ => !IsBusy && Items.Any(i => i.CanRegister));
            RegisterItemCommand        = new RelayCommand(async p => await RegisterItemAsync(p as PdfImportItem), p => !IsBusy && (p as PdfImportItem)?.CanRegister == true);
            RemoveItemCommand          = new RelayCommand(p => RemoveItem(p as PdfImportItem), p => p != null && !IsBusy);
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
                    _apiKey,
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
    }

    public class PdfImportItem : ObservableObject
    {
        public NippoData Data { get; }

        // 表示ラベル（例: p.1 / 2月27日 / 車両1603）
        public string Label => $"p.{Data.PageNumber}  {(Data.Day.HasValue ? $"{Data.Day}日" : "日付不明")}  車両{Data.VehicleNumber ?? "?"}";

        // 編集可能フィールド
        private string _day;
        public string Day { get => _day; set => SetProperty(ref _day, value); }

        private string _yuryoKm;
        public string YuryoKm { get => _yuryoKm; set => SetProperty(ref _yuryoKm, value); }

        private string _muryoKm;
        public string MuryoKm { get => _muryoKm; set => SetProperty(ref _muryoKm, value); }

        private string _shinyaMinutes;
        public string ShinyaMinutes { get => _shinyaMinutes; set => SetProperty(ref _shinyaMinutes, value); }

        private string _statusText;
        public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }

        private bool _isDone;
        public bool IsDone { get => _isDone; set { if (SetProperty(ref _isDone, value)) OnPropertyChanged(nameof(CanRegister)); } }

        public bool HasError    => string.IsNullOrEmpty(Day) || string.IsNullOrEmpty(YuryoKm) || string.IsNullOrEmpty(MuryoKm);
        public bool CanRegister => !IsDone && !string.IsNullOrEmpty(Day);

        public PdfImportItem(NippoData data)
        {
            Data          = data;
            Day           = data.Day?.ToString() ?? "";
            YuryoKm      = data.YuryoKm?.ToString() ?? "";
            MuryoKm      = data.MuryoKm?.ToString() ?? "";
            ShinyaMinutes = (data.ShinyaMinutes.HasValue && data.ShinyaMinutes > 0)
                            ? data.ShinyaMinutes.ToString() : "";

            // リトライ全失敗の場合は専用メッセージを表示
            if (data.RetryFailed)
            {
                StatusText = $"❌ 読み取り失敗（リトライ済）: {data.RetryMessage}";
                return;
            }

            var (isValid, missing) = data.ValidateRequired();
            StatusText = isValid ? "✅ 確認してください" : $"⚠️ 要確認: {missing}";
        }
    }
}
