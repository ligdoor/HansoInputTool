using System.ComponentModel;
using Newtonsoft.Json;

namespace HansoInputTool.Models
{
    /// <summary>入力タブの入力方式</summary>
    public static class TabInputMode
    {
        /// <summary>日別入力（日・距離など、日にちごとに入力する。従来の「通常」タブ）</summary>
        public const string Daily = "daily";

        /// <summary>月合計入力（総距離など、月の合計値を入力する。従来の「東日本」タブ）</summary>
        public const string Total = "total";

        public static bool IsValid(string mode) => mode == Daily || mode == Total;

        public static string DisplayName(string mode) => mode == Total ? "月合計入力" : "日別入力";
    }

    /// <summary>初期状態で用意されるタブのID（従来の「通常」「東日本」に相当）</summary>
    public static class DefaultTabIds
    {
        public const string Daily = "tab-daily";
        public const string Total = "tab-total";
    }

    /// <summary>
    /// メイン画面の入力タブ1つ分の定義。tab_settings.json に配列で保存する（配列の順番＝タブの並び順）。
    /// 画面のタブ見出しにバインドされるため、名前・入力方式の変更は画面へ通知する。
    /// </summary>
    public class TabDefinition : INotifyPropertyChanged
    {
        private string _name;
        private string _inputMode = TabInputMode.Daily;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>タブの識別子（名前を変えても変わらない。車両設定の TabId から参照される）</summary>
        public string Id { get; set; }

        /// <summary>画面に表示するタブ名（自由に変更できる）</summary>
        public string Name
        {
            get => _name;
            set
            {
                if (_name == value) return;
                _name = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
            }
        }

        /// <summary>入力方式（TabInputMode.Daily / TabInputMode.Total）</summary>
        public string InputMode
        {
            get => _inputMode;
            set
            {
                if (_inputMode == value) return;
                _inputMode = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(InputMode)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsTotalMode)));
            }
        }

        [JsonIgnore]
        public bool IsTotalMode => InputMode == TabInputMode.Total;
    }
}
