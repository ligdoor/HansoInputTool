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

        /// <summary>リスト表示用テキスト（例: "46期 4月 R7  （23件）"）</summary>
        public string DisplayText => $"{Label}  （{RecordCount}件）";
    }
}
