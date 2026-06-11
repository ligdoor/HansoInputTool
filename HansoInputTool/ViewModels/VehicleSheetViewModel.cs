using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class VehicleSheetViewModel : ObservableObject
    {
        public string OriginalSheetName { get; private set; }

        // 深夜入力方式: "time"=深夜時間（分）、"fee"=深夜料金（円）
        public List<string> 深夜入力方式リスト { get; } = new() { "time", "fee" };

        private string _lateInputMode = "time";
        public string LateInputMode
        {
            get => _lateInputMode;
            set => SetProperty(ref _lateInputMode, value);
        }

        public bool IsLateTimeModeChecked
        {
            get => LateInputMode == "time";
            set { if (value) LateInputMode = "time"; OnPropertyChanged(nameof(IsLateFeeModeChecked)); }
        }
        public bool IsLateFeeModeChecked
        {
            get => LateInputMode == "fee";
            set { if (value) LateInputMode = "fee"; OnPropertyChanged(nameof(IsLateTimeModeChecked)); }
        }

        // 「通常」を追加し、これをデフォルトとする
        public List<string> 事業所カテゴリリスト { get; } = new() { "通常", "CH富士吉田", "CH大月", "CH東富士", "東日本セレモニー" };
        public List<string> 車種リスト { get; } = new() { "寝台車", "霊柩車" };

        private string _selected事業所カテゴリ;
        public string Selected事業所カテゴリ
        {
            get => _selected事業所カテゴリ;
            set
            {
                if (SetProperty(ref _selected事業所カテゴリ, value))
                {
                    OnPropertyChanged(nameof(Is車種Visible));
                    UpdateVehicleTypeName();
                }
            }
        }

        private string _selected車種;
        public string Selected車種
        {
            get => _selected車種;
            set
            {
                if (SetProperty(ref _selected車種, value))
                {
                    UpdateVehicleTypeName();
                }
            }
        }

        private string _individualName;
        public string IndividualName
        {
            get => _individualName;
            set
            {
                if (SetProperty(ref _individualName, value))
                {
                    UpdateVehicleTypeName();
                }
            }
        }

        private string _number;
        public string Number
        {
            get => _number;
            set
            {
                if (SetProperty(ref _number, value))
                {
                    UpdateVehicleTypeName();
                }
            }
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
            Selected事業所カテゴリ = "通常"; // デフォルトを「通常」に設定
            Selected車種 = 車種リスト.First();
            UpdateVehicleTypeName(); // 初期化時に名前を生成
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

            // 「通常」以外の場合は営業所名を追加
            if (Selected事業所カテゴリ != "通常")
            {
                parts.Add(Selected事業所カテゴリ);
            }

            // 東日本セレモニー以外の場合は車種を追加
            if (Is車種Visible)
            {
                parts.Add(Selected車種);
            }

            if (!string.IsNullOrWhiteSpace(IndividualName))
            {
                parts.Add(IndividualName);
            }

            if (!string.IsNullOrWhiteSpace(Number))
            {
                parts.Add(Number);
            }

            VehicleTypeName = string.Join(" ", parts);
        }

        public void SetOriginalSheetName(string name)
        {
            OriginalSheetName = name;
        }

        private void ParseSheetName(string sheetName)
        {
            var parts = sheetName.Split(' ').ToList();

            // 営業所カテゴリを判定
            if (sheetName.Contains("東日本セレモニー"))
            {
                Selected事業所カテゴリ = "東日本セレモニー";
                parts.Remove("東日本セレモニー");
            }
            else if (sheetName.Contains("CH富士吉田"))
            {
                Selected事業所カテゴリ = "CH富士吉田";
                parts.Remove("CH富士吉田");
            }
            else if (sheetName.Contains("CH大月"))
            {
                Selected事業所カテゴリ = "CH大月";
                parts.Remove("CH大月");
            }
            else if (sheetName.Contains("CH東富士"))
            {
                Selected事業所カテゴリ = "CH東富士";
                parts.Remove("CH東富士");
            }
            else
            {
                // どの営業所名も含まれていない場合は「通常」
                Selected事業所カテゴリ = "通常";
            }

            // 車種を判定（東日本セレモニー以外の場合）
            if (Is車種Visible)
            {
                Selected車種 = 車種リスト.FirstOrDefault(s => parts.Contains(s)) ?? 車種リスト.First();
                parts.Remove(Selected車種);
            }

            // 番号を判定
            if (parts.Any() && int.TryParse(parts.Last(), out _))
            {
                Number = parts.Last();
                parts.RemoveAt(parts.Count - 1);
            }

            // 残りを個別名として扱う
            IndividualName = string.Join(" ", parts);
        }
    }
}