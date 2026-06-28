using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            string selected = ShowFolderBrowserDialog("集計ファイルが入った最上位フォルダを選択してください");
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

                    _service.ExportToExcel(data, selected, outputPath,
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

        // ---- COM フォルダダイアログ ----
        private static string ShowFolderBrowserDialog(string title)
        {
            var dialog = (IFileOpenDialog2)new FileOpenDialog2();
            try
            {
                dialog.SetOptions(0x00000020 | 0x00000040);
                dialog.SetTitle(title);
                int hr = dialog.Show(IntPtr.Zero);
                if (hr < 0) return null;
                dialog.GetResult(out IShellItem2 item);
                item.GetDisplayName(0x80058000, out string path);
                return path;
            }
            finally { Marshal.ReleaseComObject(dialog); }
        }

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialog2 { }

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog2
        {
            [PreserveSig] int Show(IntPtr p);
            void SetFileTypes(uint c, IntPtr r); void SetFileTypeIndex(uint i); void GetFileTypeIndex(out uint i);
            void Advise(IntPtr p, out uint c); void Unadvise(uint c); void SetOptions(uint f); void GetOptions(out uint f);
            void SetDefaultFolder(IShellItem2 p); void SetFolder(IShellItem2 p); void GetFolder(out IShellItem2 p);
            void GetCurrentSelection(out IShellItem2 p);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string n);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string n);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string t);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string t);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string l);
            void GetResult(out IShellItem2 p); void AddPlace(IShellItem2 p, int f);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string e);
            void Close(int h); void SetClientGuid(ref Guid g); void ClearClientData(); void SetFilter(IntPtr f);
            void GetResults(out IntPtr p); void GetSelectedItems(out IntPtr p);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem2
        {
            void BindToHandler(IntPtr p, ref Guid b, ref Guid r, out IntPtr v);
            void GetParent(out IShellItem2 p);
            void GetDisplayName(uint s, [MarshalAs(UnmanagedType.LPWStr)] out string n);
            void GetAttributes(uint m, out uint a); void Compare(IShellItem2 p, uint h, out int o);
        }
    }
}
