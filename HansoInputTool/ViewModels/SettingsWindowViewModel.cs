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

            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(
                JsonConvert.SerializeObject(currentRates));

            var currentSheets = _excelHandler.GetVehicleSheetNames();
            VehicleSheetList = new ObservableCollection<VehicleSheetViewModel>(
                currentSheets.Select(s => new VehicleSheetViewModel(s)));

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

            var sheetName = !string.IsNullOrEmpty(SelectedVehicle.VehicleTypeName)
                ? SelectedVehicle.VehicleTypeName
                : "新しい車両";

            var result = MessageBox.Show(
                $"車両 '{sheetName}' を削除しますか？\nExcelシートも同時に削除されます。",
                "削除確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

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
                MessageBox.Show("車両名が空の項目があります。\nカテゴリ、個別名、番号などを入力してください。",
                    "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var duplicate = VehicleSheetList.GroupBy(v => v.VehicleTypeName)
                .FirstOrDefault(g => g.Count() > 1);
            if (duplicate != null)
            {
                MessageBox.Show($"車両名 '{duplicate.Key}' が重複しています。\n個別名や番号を変更してください。",
                    "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var originalSheetNames = _excelHandler.GetVehicleSheetNames();
                var finalSheetVMs = VehicleSheetList.ToList();

                var deletedSheetNames = originalSheetNames
                    .Where(origName => !finalSheetVMs.Any(vm => vm.OriginalSheetName == origName))
                    .ToList();

                var renameMap = finalSheetVMs
                    .Where(vm => vm.OriginalSheetName != null && vm.OriginalSheetName != vm.VehicleTypeName)
                    .ToDictionary(vm => vm.OriginalSheetName, vm => vm.VehicleTypeName);

                var addedSheets = new List<(string newName, string templateName)>();
                foreach (var vm in finalSheetVMs.Where(vm => vm.OriginalSheetName == null))
                {
                    string templateSheetName = FindTemplateSheetName(vm.Selected事業所カテゴリ, vm.Selected車種);
                    if (string.IsNullOrEmpty(templateSheetName) || !_excelHandler.SheetNames.Contains(templateSheetName))
                    {
                        MessageBox.Show($"カテゴリ '{vm.Selected事業所カテゴリ}' のテンプレートシート '{templateSheetName}' が見つかりません。",
                            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    addedSheets.Add((vm.VehicleTypeName, templateSheetName));
                }

                _excelHandler.SyncAllVehicleSheets(deletedSheetNames, renameMap, addedSheets);

                string ratesJson = JsonConvert.SerializeObject(Rates, Formatting.Indented);
                File.WriteAllText(_ratesFilePath, ratesJson);

                _mainViewModel.UpdateRatesAndReload(Rates);
                MessageBox.Show("設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);

                if (parameter is Window window)
                {
                    window.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"設定の保存中にエラーが発生しました。\n{ex.Message}",
                    "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string FindTemplateSheetName(string category, string shashu)
        {
            var sheetNames = _excelHandler.GetVehicleSheetNames();

            if (category == "CH大月")
                return sheetNames.FirstOrDefault(s => s.Contains("CH大月")) ?? "寝台車 30";

            if (category == "CH東富士")
                return sheetNames.FirstOrDefault(s => s.Contains("CH東富士")) ?? "寝台車 30";

            if (category == "東日本セレモニー")
                return "東日本セレモニー 1961";

            if (category == "CH富士吉田" || category == "通常")
            {
                if (shashu == "寝台車") return "寝台車 30";
                if (shashu == "霊柩車") return sheetNames.FirstOrDefault(s => s.Contains("霊柩車"));
            }

            return "寝台車 30";
        }
    }
}
