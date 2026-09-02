using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class EditWindowViewModel : ObservableObject
    {
        private readonly MainViewModel _mainViewModel;
        private readonly string _sheetName;
        private readonly int _rowIndex;   // Excel使用時の行番号
        private readonly long _dbId;       // DB使用時の主キー

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

        /// <summary>動的フラグチェックボックスのリスト</summary>
        public ObservableCollection<FlagCheckBoxItem> FlagItems { get; } = new();

        /// <summary>この車両が給油管理表への記録対象かどうか</summary>
        public bool IsFuelTrackedVehicle { get; }

        /// <summary>この日（行）に紐づく給油記録の一覧（編集・削除・追加が可能）</summary>
        public ObservableCollection<EditableFuelItem> FuelEntries { get; } = new();

        public ICommand SaveCommand { get; }
        public ICommand AddFuelEntryCommand { get; }

        public EditWindowViewModel(MainViewModel mainViewModel, string sheetName, RowData rowData)
        {
            _mainViewModel = mainViewModel;
            _sheetName     = sheetName;
            _rowIndex      = rowData.RowIndex;
            _dbId          = rowData.DbId;   // DB使用時の主キー

            IsOotsukiSheet = mainViewModel.IsFeeMode(sheetName);
            WindowTitle    = $"行 {rowData.B_Day}日 を編集 - {sheetName}";

            Day       = rowData.B_Day?.ToString();
            YuryoKm   = rowData.D_YuryoKm?.ToString();
            MuryoKm   = rowData.E_MuryoKm?.ToString();
            LateValue = IsOotsukiSheet
                ? rowData.H_LateFeeOotsuki?.ToString()
                : rowData.K_LateMinutes?.ToString();

            // 動的フラグをRowDataから復元
            var flagService = mainViewModel.FlagService;
            if (flagService != null)
            {
                foreach (var flag in flagService.Flags.OrderBy(f => f.Order))
                {
                    var item = new FlagCheckBoxItem(flag)
                    {
                        IsChecked = rowData.GetFlag(flag.Id)
                    };
                    FlagItems.Add(item);
                }
            }

            // 給油記録の読み込み（給油管理対象車両のみ）。この行(DbId)に紐付いた記録を優先し、
            // 無ければ日付だけで一致する古いデータをフォールバック表示する。
            IsFuelTrackedVehicle = mainViewModel.IsFuelTracked(sheetName);
            if (IsFuelTrackedVehicle && rowData.B_Day.HasValue)
            {
                foreach (var fuel in mainViewModel.GetFuelRecordsForRow(sheetName, _dbId, rowData.B_Day.Value))
                    FuelEntries.Add(CreateFuelItem(fuel.Id, fuel.OdometerKm.ToString(), fuel.Liters.ToString()));
            }

            SaveCommand = new RelayCommand(SaveEdit);
            AddFuelEntryCommand = new RelayCommand(_ => FuelEntries.Add(CreateFuelItem(0, "", "")));
        }

        private EditableFuelItem CreateFuelItem(long id, string km, string liters)
        {
            var item = new EditableFuelItem { Id = id, OdometerKm = km, Liters = liters };
            item.DeleteCommand = new RelayCommand(_ => RemoveFuelItem(item));
            return item;
        }

        private void RemoveFuelItem(EditableFuelItem item)
        {
            if (item.Id > 0)
            {
                if (MessageBox.Show("この給油記録を削除しますか？\nこの操作は元に戻せません。",
                        "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    return;

                try
                {
                    _mainViewModel.DeleteFuelRecord(_sheetName, item.Id);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show(ex.Message, "削除できません", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            FuelEntries.Remove(item);
        }

        private void SaveEdit(object parameter)
        {
            if (string.IsNullOrWhiteSpace(Day))
            {
                MessageBox.Show("日付は必須です。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var values = new Dictionary<string, double?>();

            if (!TryParseValue(Day,      "日(B)",      out var dayVal))     return;
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

            // 給油欄の入力チェック（新規で追加したが未入力のままの行は無視、既存行や入力済みの新規行はエラーチェック）
            var fuelToSave = new List<(EditableFuelItem item, double km, double liters)>();
            if (IsFuelTrackedVehicle)
            {
                foreach (var entry in FuelEntries)
                {
                    entry.ErrorText = string.Empty;
                    bool kmEmpty     = string.IsNullOrWhiteSpace(entry.OdometerKm);
                    bool litersEmpty = string.IsNullOrWhiteSpace(entry.Liters);

                    if (entry.Id == 0 && kmEmpty && litersEmpty)
                        continue; // 追加したが未入力のままの行はスキップ

                    if (!double.TryParse(entry.OdometerKm, out var km) || km <= 0)
                    {
                        entry.ErrorText = "給油時Kmを正しく入力してください。";
                        return;
                    }
                    if (!double.TryParse(entry.Liters, out var liters) || liters <= 0)
                    {
                        entry.ErrorText = "給油㍑数を正しく入力してください。";
                        return;
                    }
                    fuelToSave.Add((entry, km, liters));
                }
            }

            var flagStates = FlagItems.ToDictionary(f => f.Id, f => f.IsChecked);
            // DB使用時は DbId を、Excel使用時は RowIndex を渡す
            int idToPass = (_dbId > 0) ? (int)_dbId : _rowIndex;
            try
            {
                _mainViewModel.UpdateRowData(_sheetName, idToPass, values, flagStates);

                // 給油記録の保存（日付が変更されていれば新しい日に紐づけ直す。この行のDbIdに紐付ける）
                int fuelDay = (int)dayVal.Value;
                foreach (var (entry, km, liters) in fuelToSave)
                {
                    if (entry.Id > 0)
                        _mainViewModel.UpdateFuelRecord(_sheetName, entry.Id, _dbId, fuelDay, km, liters);
                    else
                        _mainViewModel.AddFuelRecord(_sheetName, _dbId, fuelDay, km, liters);
                }

                ((Window)parameter).Close();
            }
            catch (InvalidOperationException ex)
            {
                // 確定済みセッションのデータを更新しようとした場合など、意図的にブロックしている操作
                MessageBox.Show(ex.Message, "更新できません", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static bool TryParseValue(string input, string fieldName, out double? result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input)) return true;
            if (double.TryParse(input, out double parsedValue)) { result = parsedValue; return true; }
            MessageBox.Show($"「{input}」は {fieldName} の数値として認識できません。",
                "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }
}
