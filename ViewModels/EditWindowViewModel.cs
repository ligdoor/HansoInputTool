using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class EditWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly string _sheetName;
        private readonly int _rowIndex;

        public string WindowTitle { get; }
        public bool IsOotsukiSheet { get; }

        private string _day;
        public string Day { get => _day; set => SetProperty(ref _day, value); }

        private string _yuryoKm;
        public string YuryoKm { get => _yuryoKm; set => SetProperty(ref _yuryoKm, value); }

        private string _muryoKm;
        public string MuryoKm { get => _muryoKm; set => SetProperty(ref _muryoKm, value); }

        private string _lateValue;
        public string LateValue { get => _lateValue; set => SetProperty(ref _lateValue, value); }

        private bool _isKoryo;
        public bool IsKoryo { get => _isKoryo; set => SetProperty(ref _isKoryo, value); }

        public ICommand SaveCommand { get; }

        public EditWindowViewModel(MainViewModel mainViewModel, string sheetName, RowData rowData)
        {
            _mainViewModel = mainViewModel;
            _sheetName = sheetName;
            _rowIndex = rowData.RowIndex;

            IsOotsukiSheet = sheetName.Contains("大月");
            WindowTitle = $"行 {rowData.RowIndex} を編集 - {sheetName}";

            Day = rowData.B_Day?.ToString();
            YuryoKm = rowData.D_YuryoKm?.ToString();
            MuryoKm = rowData.E_MuryoKm?.ToString();
            LateValue = IsOotsukiSheet ? rowData.H_LateFeeOotsuki?.ToString() : rowData.K_LateMinutes?.ToString();
            IsKoryo = rowData.L_IsKoryo == 1;

            SaveCommand = new RelayCommand(SaveEdit);
        }

        private void SaveEdit(object parameter)
        {
            if (string.IsNullOrWhiteSpace(Day))
            {
                MessageBox.Show("日付は必須です。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var values = new Dictionary<string, double?>();

            if (!TryParseValue(Day, "日(B)", out var dayVal)) return;
            values["日(B)"] = dayVal;

            if (!TryParseValue(YuryoKm, "有料キロ(D)", out var yuryoKmVal)) return;
            values["有料キロ(D)"] = yuryoKmVal;

            if (!TryParseValue(MuryoKm, "無料キロ(E)", out var muryoKmVal)) return;
            values["無料キロ(E)"] = muryoKmVal;

            if (IsOotsukiSheet)
            {
                if (!TryParseValue(LateValue, "深夜料金(H)", out var lateVal)) return;
                values["深夜料金(H)"] = lateVal;
            }
            else
            {
                if (!TryParseValue(LateValue, "深夜時間(K)", out var lateVal)) return;
                values["深夜時間(K)"] = lateVal;
            }

            _mainViewModel.UpdateRowData(_sheetName, _rowIndex, values, IsKoryo);
            ((Window)parameter).Close();
        }

        private static bool TryParseValue(string input, string fieldName, out double? result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input))
            {
                return true; // Empty is allowed for non-required fields
            }

            if (double.TryParse(input, out double parsedValue))
            {
                result = parsedValue;
                return true;
            }

            MessageBox.Show($"「{input}」は {fieldName} の数値として認識できません。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }
}