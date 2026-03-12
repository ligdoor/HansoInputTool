using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using Microsoft.Win32;

namespace HansoInputTool.ViewModels
{
    public class VehicleAnnualSummaryViewModel : ObservableObject
    {
        private readonly VehicleAnnualSummaryService _service = new();

        private string _folderPath = "";
        public string FolderPath
        {
            get => _folderPath;
            set => SetProperty(ref _folderPath, value);
        }

        private int _startYear = DateTime.Today.Month >= 5 ? DateTime.Today.Year : DateTime.Today.Year - 1;
        public int StartYear
        {
            get => _startYear;
            set => SetProperty(ref _startYear, value);
        }

        private int _startMonth = 5;
        public int StartMonth
        {
            get => _startMonth;
            set => SetProperty(ref _startMonth, value);
        }

        private int _endYear = DateTime.Today.Month >= 5 ? DateTime.Today.Year + 1 : DateTime.Today.Year;
        public int EndYear
        {
            get => _endYear;
            set => SetProperty(ref _endYear, value);
        }

        private int _endMonth = 4;
        public int EndMonth
        {
            get => _endMonth;
            set => SetProperty(ref _endMonth, value);
        }

        private string _statusMessage = "フォルダを選択して実行してください。";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (SetProperty(ref _isBusy, value))
                    OnPropertyChanged(nameof(IsNotBusy));
            }
        }
        public bool IsNotBusy => !_isBusy;

        public ICommand SelectFolderCommand { get; }
        public ICommand ExecuteCommand { get; }

        public VehicleAnnualSummaryViewModel()
        {
            SelectFolderCommand = new RelayCommand(_ => SelectFolder());
            ExecuteCommand = new RelayCommand(_ => Execute(),
                                                   _ => IsNotBusy && !string.IsNullOrEmpty(FolderPath));
        }

        private void SelectFolder()
        {
            // WPF標準のOpenFileDialogを使ってフォルダ選択
            // （UseWindowsFormsが不要なためクラッシュしない）
            var dialog = new OpenFileDialog
            {
                Title = "集計ファイルが入った最上位フォルダを選択（フォルダ内の任意のファイルを選択してください）",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "フォルダを選択",
                Filter = "フォルダ|*.",
                ValidateNames = false
            };

            if (dialog.ShowDialog() == true)
            {
                string selectedFolder = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(selectedFolder))
                {
                    FolderPath = selectedFolder;
                    StatusMessage = $"フォルダ選択済み: {FolderPath}";
                }
            }
        }

        private async void Execute()
        {
            if (string.IsNullOrEmpty(FolderPath))
            {
                MessageBox.Show("フォルダを選択してください。", "確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (StartYear * 100 + StartMonth > EndYear * 100 + EndMonth)
            {
                MessageBox.Show("終了年月は開始年月より後にしてください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title = "出力先を選択",
                Filter = "Excel ファイル (*.xlsx)|*.xlsx",
                FileName = $"運輸実績_{StartYear}年{StartMonth}月-{EndYear}年{EndMonth}月.xlsx",
                InitialDirectory = FolderPath
            };
            if (saveDialog.ShowDialog() != true) return;

            string outputPath = saveDialog.FileName;

            IsBusy = true;
            StatusMessage = "集計中...";

            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    var data = _service.LoadData(
                        FolderPath, StartYear, StartMonth, EndYear, EndMonth);

                    if (data.Count == 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            StatusMessage = "対象ファイルが見つかりませんでした。";
                            MessageBox.Show(
                                "指定した期間のファイルが見つかりませんでした。\nフォルダとファイル名を確認してください。",
                                "データなし", MessageBoxButton.OK, MessageBoxImage.Warning);
                        });
                        return;
                    }

                    _service.ExportToExcel(data, outputPath,
                        StartYear, StartMonth, EndYear, EndMonth);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        StatusMessage = $"完了！ → {Path.GetFileName(outputPath)}";
                        if (MessageBox.Show(
                            $"集計が完了しました。\nファイルを開きますか？\n\n{outputPath}",
                            "完了", MessageBoxButton.YesNo, MessageBoxImage.Information)
                            == MessageBoxResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = outputPath,
                                UseShellExecute = true
                            });
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
                MessageBox.Show($"エラーが発生しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }
    }
}