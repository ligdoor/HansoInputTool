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

        /// <summary>
        /// 「保存」ボタンで保護された（保存済み）かどうか。
        /// 保存済みの月は、転記終了時の自動クリアや「クリア」でデータが消えない（編集は自由にできる）。
        /// </summary>
        public bool   IsSaved     { get; set; }
        /// <summary>保存した日時（未保存の場合はnull）</summary>
        public string SavedAt     { get; set; }

        /// <summary>リスト表示用テキスト（例: "46期 4月 R7  （23件） 💾保存済 🔒確定済"）</summary>
        public string DisplayText =>
            $"{Label}  （{RecordCount}件）"
            + (IsSaved ? " 💾保存済" : "")
            + (IsConfirmed ? " 🔒確定済" : "");
    }
}
