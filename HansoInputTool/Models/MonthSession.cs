namespace HansoInputTool.Models
{
    /// <summary>month_sessions テーブルの1行を表すモデル</summary>
    public class MonthSession
    {
        public long   Id          { get; set; }
        public string Period      { get; set; }
        public string Month       { get; set; }
        public string RNumber     { get; set; }
        public string Label       { get; set; }
        public string CreatedAt   { get; set; }
        public int    RecordCount { get; set; }

        /// <summary>確定済みかどうか（確定済みは誤操作での編集・削除から保護される）</summary>
        public bool   IsConfirmed { get; set; }
        /// <summary>確定した日時（未確定の場合はnull）</summary>
        public string ConfirmedAt { get; set; }

        /// <summary>リスト表示用テキスト（例: "46期 4月 R7  （23件） 🔒確定済"）</summary>
        public string DisplayText => IsConfirmed
            ? $"{Label}  （{RecordCount}件） 🔒確定済"
            : $"{Label}  （{RecordCount}件）";
    }
}
