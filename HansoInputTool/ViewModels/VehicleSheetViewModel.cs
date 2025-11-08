using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class VehicleSheetViewModel : ObservableObject
    {
        public string OriginalSheetName { get; private set; }

        public List<string> 事業所カテゴリリスト { get; } = new() { "CH富士吉田", "CH大月", "CH東富士", "東日本セレモニー" };
        public List<string> 車種リスト { get; } = new() { "寝台車", "霊柩車" };

        private string _selected事業所カテゴリ;
        public string Selected事業所カテゴリ
        {
            get => _selected事業所カテゴリ;
            set { if (SetProperty(ref _selected事業所カテゴリ, value)) { OnPropertyChanged(nameof(Is車種Visible)); UpdateVehicleTypeName(); } }
        }

        private string _selected車種;
        public string Selected車種
        {
            get => _selected車種;
            set { if (SetProperty(ref _selected車種, value)) { UpdateVehicleTypeName(); } }
        }

        private string _individualName;
        public string IndividualName
        {
            get => _individualName;
            set { if (SetProperty(ref _individualName, value)) { UpdateVehicleTypeName(); } }
        }

        private string _number;
        public string Number
        {
            get => _number;
            set { if (SetProperty(ref _number, value)) { UpdateVehicleTypeName(); } }
        }

        private string _vehicleTypeName;
        public string VehicleTypeName
        {
            get => _vehicleTypeName;
            private set => SetProperty(ref _vehicleTypeName, value);
        }

        public bool Is車種Visible => Selected事業所カテゴリ != "東日本セレモニー";

        public VehicleSheetViewModel()
        {
            OriginalSheetName = null;
            Selected事業所カテゴリ = 事業所カテゴリリスト.First();
            Selected車種 = 車種リスト.First();
        }

        public VehicleSheetViewModel(string sheetName)
        {
            OriginalSheetName = sheetName;
            VehicleTypeName = sheetName;
            ParseSheetName(sheetName);
        }

        private void UpdateVehicleTypeName()
        {
            var parts = new List<string>();
            if (Selected事業所カテゴリ != "CH富士吉田") parts.Add(Selected事業所カテゴリ);
            if (Is車種Visible) parts.Add(Selected車種);
            if (!string.IsNullOrWhiteSpace(IndividualName)) parts.Add(IndividualName);
            if (!string.IsNullOrWhiteSpace(Number)) parts.Add(Number);
            VehicleTypeName = string.Join(" ", parts);
        }

        public void SetOriginalSheetName(string name)
        {
            OriginalSheetName = name;
        }

        private void ParseSheetName(string sheetName)
        {
            var parts = sheetName.Split(' ').ToList();
            Selected事業所カテゴリ = 事業所カテゴリリスト.FirstOrDefault(c => sheetName.Contains(c) && c != "CH富士吉田") ?? "CH富士吉田";
            if (Selected事業所カテゴリ != "CH富士吉田") parts.Remove(Selected事業所カテゴリ);
            if (Is車種Visible)
            {
                Selected車種 = 車種リスト.FirstOrDefault(s => parts.Contains(s)) ?? 車種リスト.First();
                parts.Remove(Selected車種);
            }
            if (parts.Any() && int.TryParse(parts.Last(), out _))
            {
                Number = parts.Last();
                parts.RemoveAt(parts.Count - 1);
            }
            IndividualName = string.Join(" ", parts);
        }
    }
}