using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.ViewModels.Base;
using Newtonsoft.Json;

namespace HansoInputTool.ViewModels
{
    public class SettingsWindowViewModel : ObservableObject
    {
        private readonly string _ratesFilePath;
        private readonly string _columnMapFilePath;
        private readonly MainViewModel _mainViewModel;

        public Dictionary<string, RateInfo> Rates { get; set; }
        public ColumnMapping ColumnMap { get; set; }

        public ICommand SaveCommand { get; }
        public ICommand CancelCommand { get; }

        public SettingsWindowViewModel(
            Dictionary<string, RateInfo> currentRates,
            ColumnMapping currentColumnMap,
            string ratesFilePath,
            string columnMapFilePath,
            MainViewModel mainViewModel)
        {
            // Deep copy to avoid modifying original data until save
            Rates = JsonConvert.DeserializeObject<Dictionary<string, RateInfo>>(JsonConvert.SerializeObject(currentRates));
            ColumnMap = JsonConvert.DeserializeObject<ColumnMapping>(JsonConvert.SerializeObject(currentColumnMap));

            _ratesFilePath = ratesFilePath;
            _columnMapFilePath = columnMapFilePath;
            _mainViewModel = mainViewModel;

            SaveCommand = new RelayCommand(SaveSettings);
            CancelCommand = new RelayCommand(p => ((Window)p).Close());
        }

        private void SaveSettings(object parameter)
        {
            try
            {
                // 料金設定を保存
                string ratesJson = JsonConvert.SerializeObject(Rates, Formatting.Indented);
                File.WriteAllText(_ratesFilePath, ratesJson);

                // 列マッピング設定を保存
                string columnMapJson = JsonConvert.SerializeObject(ColumnMap, Formatting.Indented);
                File.WriteAllText(_columnMapFilePath, columnMapJson);

                // MainViewModelのデータを更新
                _mainViewModel.Rates = Rates;
                _mainViewModel.UpdateColumnMap(ColumnMap); // MainViewModelに新しいメソッドを追加

                MessageBox.Show("設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);

                ((Window)parameter).Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"設定の保存に失敗しました。\n{ex.Message}", "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}