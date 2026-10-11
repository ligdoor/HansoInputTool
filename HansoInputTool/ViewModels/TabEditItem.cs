using HansoInputTool.Models;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>設定画面の「入力タブ」一覧で編集する、タブ1つ分の編集用データ</summary>
    public class TabEditItem : ObservableObject
    {
        public string Id { get; }

        public TabEditItem(string id, string name, string inputMode)
        {
            Id = id;
            _name = name;
            _inputMode = TabInputMode.IsValid(inputMode) ? inputMode : TabInputMode.Daily;
        }

        private string _name;
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        private string _inputMode;
        public string InputMode
        {
            get => _inputMode;
            set
            {
                if (!SetProperty(ref _inputMode, value)) return;
                OnPropertyChanged(nameof(IsDailyChecked));
                OnPropertyChanged(nameof(IsTotalChecked));
                OnPropertyChanged(nameof(IsTotalMode));
                OnPropertyChanged(nameof(InputModeDisplay));
            }
        }

        public bool IsTotalMode => InputMode == TabInputMode.Total;
        public string InputModeDisplay => TabInputMode.DisplayName(InputMode);

        public bool IsDailyChecked
        {
            get => InputMode == TabInputMode.Daily;
            set { if (value) InputMode = TabInputMode.Daily; }
        }

        public bool IsTotalChecked
        {
            get => InputMode == TabInputMode.Total;
            set { if (value) InputMode = TabInputMode.Total; }
        }

        private int _vehicleCount;
        /// <summary>このタブに所属している車両の台数（画面表示・入力方式変更の可否判定用）</summary>
        public int VehicleCount
        {
            get => _vehicleCount;
            set
            {
                if (!SetProperty(ref _vehicleCount, value)) return;
                OnPropertyChanged(nameof(CanChangeInputMode));
            }
        }

        /// <summary>車両が所属していないタブだけ、入力方式を変更できる（シートのレイアウトが違うため）</summary>
        public bool CanChangeInputMode => VehicleCount == 0;
    }
}
