using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
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

        // [デザイン維持対応] 年間集計のひな形（annual_results.xlsx）は他の設定ファイルと同じく
        // HansoDataフォルダに配置する運用とする。
        private static string AnnualTemplateFilePath =>
            Path.Combine(
                App.DataPath ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data"),
                "annual_results.xlsx");

        // ── フォルダ・期間 ──
        private string _folderPath = "";
        public string FolderPath
        {
            get => _folderPath;
            set => SetProperty(ref _folderPath, value);
        }

        private int _startYear = DateTime.Today.Month >= 5
            ? DateTime.Today.Year : DateTime.Today.Year - 1;
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

        private int _endYear = DateTime.Today.Month >= 5
            ? DateTime.Today.Year + 1 : DateTime.Today.Year;
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

        // ── 車両チェックリスト ──
        public ObservableCollection<VehicleEntryViewModel> Vehicles { get; } = new();

        private bool _hasVehicles;
        public bool HasVehicles
        {
            get => _hasVehicles;
            set => SetProperty(ref _hasVehicles, value);
        }

        // ── ステータス ──
        private string _statusMessage = "フォルダを選択して「車両を読み込む」を押してください。";
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

        // ── コマンド ──
        public ICommand SelectFolderCommand    { get; }
        public ICommand ScanVehiclesCommand    { get; }
        public ICommand SelectAllCommand       { get; }
        public ICommand DeselectAllCommand     { get; }
        public ICommand ExecuteCommand         { get; }

        public VehicleAnnualSummaryViewModel()
        {
            SelectFolderCommand = new RelayCommand(_ => SelectFolder());
            ScanVehiclesCommand = new RelayCommand(
                _ => ScanVehicles(),
                _ => IsNotBusy && !string.IsNullOrEmpty(FolderPath));
            SelectAllCommand    = new RelayCommand(
                _ => { foreach (var v in Vehicles) v.IsChecked = true; },
                _ => Vehicles.Count > 0);
            DeselectAllCommand  = new RelayCommand(
                _ => { foreach (var v in Vehicles) v.IsChecked = false; },
                _ => Vehicles.Count > 0);
            ExecuteCommand = new RelayCommand(
                _ => Execute(),
                _ => IsNotBusy && Vehicles.Any(v => v.IsChecked));
        }

        private void SelectFolder()
        {
            string selected = FolderPickerService.ShowFolderBrowserDialog("集計ファイルが入った最上位フォルダを選択してください");
            if (!string.IsNullOrEmpty(selected))
            {
                FolderPath = selected;
                Vehicles.Clear();
                HasVehicles = false;
                StatusMessage = $"フォルダ選択済み: {FolderPath}　←「車両を読み込む」を押してください";
            }
        }

        private async void ScanVehicles()
        {
            if (StartYear * 100 + StartMonth > EndYear * 100 + EndMonth)
            {
                MessageBox.Show("終了年月は開始年月より後にしてください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IsBusy = true;
            StatusMessage = "Excelファイルを読み込んで車両リストを取得中...";
            Vehicles.Clear();
            HasVehicles = false;

            try
            {
                var entries = await System.Threading.Tasks.Task.Run(() =>
                    _service.ScanVehicles(FolderPath,
                        StartYear, StartMonth, EndYear, EndMonth));

                foreach (var e in entries)
                    Vehicles.Add(new VehicleEntryViewModel(e));

                HasVehicles = Vehicles.Count > 0;

                if (Vehicles.Count == 0)
                    StatusMessage = "対象期間のファイルが見つかりませんでした。";
                else
                {
                    int unknown = Vehicles.Count(v => !v.IsKnown);
                    string unknownNote = unknown > 0
                        ? $"（うち未分類: {unknown}件）" : "";
                    StatusMessage =
                        $"{Vehicles.Count}台の車両を検出しました{unknownNote}。集計したい車両にチェックを入れて「集計を実行」を押してください。";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
                MessageBox.Show($"車両の読み込み中にエラーが発生しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async void Execute()
        {
            var selected = Vehicles
                .Where(v => v.IsChecked)
                .Select(v => { v.Entry.IsChecked = true; return v.Entry; })
                .ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show("集計する車両を1台以上選択してください。", "確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title  = "出力先を選択",
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
                        FolderPath,
                        StartYear, StartMonth, EndYear, EndMonth,
                        selected);

                    if (data.Count == 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            StatusMessage = "選択した車両のデータが見つかりませんでした。";
                            MessageBox.Show(
                                "選択した車両の実績データが見つかりませんでした。\n車両の選択やフォルダを確認してください。",
                                "データなし", MessageBoxButton.OK, MessageBoxImage.Warning);
                        });
                        return;
                    }

                    if (!File.Exists(AnnualTemplateFilePath))
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            StatusMessage = "ひな形ファイルが見つかりません。";
                            MessageBox.Show(
                                $"年間集計のひな形ファイルが見つかりません。\n\n{AnnualTemplateFilePath}\n\nHansoDataフォルダにannual_results.xlsxを配置してください。",
                                "ひな形ファイルなし", MessageBoxButton.OK, MessageBoxImage.Warning);
                        });
                        return;
                    }

                    _service.ExportToExcel(data, selected, outputPath,
                        StartYear, StartMonth, EndYear, EndMonth, AnnualTemplateFilePath);

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
                                FileName = outputPath, UseShellExecute = true
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
