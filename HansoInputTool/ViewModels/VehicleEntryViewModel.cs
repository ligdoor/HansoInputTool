using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// チェックリスト用ラッパー（INotifyPropertyChanged対応）
    /// </summary>
    public class VehicleEntryViewModel : ObservableObject
    {
        private bool _isChecked;
        public VehicleEntry Entry { get; }

        public string Label   => Entry.Label;
        public bool   IsKnown => Entry.IsKnown;

        public bool IsChecked
        {
            get => _isChecked;
            set => SetProperty(ref _isChecked, value);
        }

        public VehicleEntryViewModel(VehicleEntry entry)
        {
            Entry      = entry;
            _isChecked = entry.IsChecked;
        }
    }
}
