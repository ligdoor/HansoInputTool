using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using Newtonsoft.Json;

namespace HansoInputTool.ViewModels
{
    public class SettingsWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly ExcelHandler _excelHandler;
        private readonly string _ratesFilePath;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ObservableCollection<VehicleSheetViewModel> VehicleSheetList { get; set; }

        private VehicleSheetViewModel _selectedVehicle;
        public VehicleSheetViewModel SelectedVehicle
        {
            get => _selectedVehicle;
            set => SetProperty(ref _selectedVehicle, value);
        }

        public ICommand AddVehicleCommand { get; }
        public ICommand DeleteVehicleCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }

        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ExcelHandler excelHandler,
            string ratesFilePath,
            MainViewModel mainViewModel)
        {
            _excelHandler = excelHandler;
            _ratesFilePath = ratesFilePath;
            _mainViewModel = mainViewModel;

            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(JsonConvert.SerializeObject(currentRates));
            var currentSheets = _excelHandler.GetVehicleSheetNames();
            VehicleSheetList = new ObservableCollection<VehicleSheetViewModel>(
                currentSheets.Select(s => new VehicleSheetViewModel(s))
            );

            AddVehicleCommand = new RelayCommand(p => AddVehicle());
            DeleteVehicleCommand = new RelayCommand(p => DeleteVehicle(), p => SelectedVehicle != null);
            SaveCommand = new RelayCommand(p => SaveSettings(p));
            CancelCommand = new RelayCommand(p => ((Window)p).Close());
        }

        private void AddVehicle()
        {
            var newVehicle = new VehicleSheetViewModel();
            VehicleSheetList.Add(newVehicle);
            SelectedVehicle = newVehicle;
        }

        private void DeleteVehicle()
        {
            if (SelectedVehicle == null) return;
            var sheetName = !string.IsNullOrEmpty(SelectedVehicle.VehicleTypeName) ? SelectedVehicle.VehicleTypeName : "新しい車両";
            var result = MessageBox.Show($"車両 '{sheetName}' を削除しますか？\nExcelシートも同時に削除されます。", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                VehicleSheetList.Remove(SelectedVehicle);
                SelectedVehicle = null;
            }
        }

        private void SaveSettings(object parameter)
        {
            if (VehicleSheetList.Any(v => string.IsNullOrWhiteSpace(v.VehicleTypeName)))
            {
                MessageBox.Show("車両名が空の項目があります。\nカテゴリ、個別名、番号などを入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var duplicate = VehicleSheetList.GroupBy(v => v.VehicleTypeName)
                                        .FirstOrDefault(g => g.Count() > 1);
            if (duplicate != null)
            {
                MessageBox.Show($"車両名 '{duplicate.Key}' が重複しています。\n個別名や番号を変更してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var existingSheets = _excelHandler.GetVehicleSheetNames();
                var uiVehicles = VehicleSheetList.ToList();
                var sheetsToDelete = existingSheets.Except(uiVehicles.Select(v => v.VehicleTypeName)).ToList();

                foreach (var sheetName in sheetsToDelete)
                {
                    _excelHandler.DeleteVehicleSheet(sheetName);
                }

                foreach (var vehicleVM in uiVehicles)
                {
                    if (vehicleVM.OriginalSheetName == null) // 新規追加
                    {
                        string templateSheetName = FindTemplateSheetName(vehicleVM.Selected事業所カテゴリ, vehicleVM.Selected車種);
                        if (string.IsNullOrEmpty(templateSheetName))
                        {
                            MessageBox.Show($"カテゴリ '{vehicleVM.Selected事業所カテゴリ}' のコピー元となるテンプレートシートが見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        _excelHandler.AddVehicleSheet(vehicleVM.VehicleTypeName, templateSheetName);
                    }
                    else if (vehicleVM.OriginalSheetName != vehicleVM.VehicleTypeName) // 名前変更
                    {
                        _excelHandler.RenameVehicleSheet(vehicleVM.OriginalSheetName, vehicleVM.VehicleTypeName);
                    }
                }

                // 月間集計シートの同期
                _excelHandler.UpdateMonthlySummarySheet(uiVehicles.Select(v => v.VehicleTypeName).ToList());

                _excelHandler.Save();

                string json = JsonConvert.SerializeObject(Rates, Formatting.Indented);
                File.WriteAllText(_ratesFilePath, json);

                _mainViewModel.UpdateRatesAndReload(Rates);

                MessageBox.Show("設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);

                if (parameter is Window window)
                {
                    window.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"設定の保存中にエラーが発生しました。\n{ex.Message}", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string FindTemplateSheetName(string category, string shashu)
        {
            var sheetNames = _excelHandler.GetVehicleSheetNames();

            if (category == "CH大月") return sheetNames.FirstOrDefault(s => s.Contains("CH大月"));
            if (category == "CH東富士") return sheetNames.FirstOrDefault(s => s.Contains("CH東富士"));
            if (category == "東日本セレモニー") return "東日本セレモニー 1961";

            if (category == "通常")
            {
                if (shashu == "寝台車") return "寝台車 30";
                if (shashu == "霊柩車") return sheetNames.FirstOrDefault(s => s.Contains("霊柩車"));
            }

            return "寝台車 30"; // デフォルト
        }
    }
}