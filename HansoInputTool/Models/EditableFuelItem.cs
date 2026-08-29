using System.Windows.Input;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.Models
{
    /// <summary>
    /// 編集画面（EditWindow）で1件の給油記録を表す編集用アイテム。
    /// Id が 0 のときは「まだDBに保存されていない新規入力行」を表す。
    /// </summary>
    public class EditableFuelItem : ObservableObject
    {
        /// <summary>DB上のID。0なら新規（未保存）</summary>
        public long Id { get; set; }

        private string _odometerKm;
        public string OdometerKm { get => _odometerKm; set => SetProperty(ref _odometerKm, value); }

        private string _liters;
        public string Liters { get => _liters; set => SetProperty(ref _liters, value); }

        private string _errorText;
        public string ErrorText { get => _errorText; set => SetProperty(ref _errorText, value); }

        public ICommand DeleteCommand { get; set; }
    }
}
