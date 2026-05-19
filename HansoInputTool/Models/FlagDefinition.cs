using System.Collections.Generic;

namespace HansoInputTool.Models
{
    /// <summary>
    /// チェックボックスフラグの種類
    /// </summary>
    public enum FlagType
    {
        CountOnly,    // 回数のみ（エンバーミングタイプ）
        WithAmount    // 金額あり（行旅死亡人タイプ）
    }

    /// <summary>
    /// 金額ありタイプの金額計算方法
    /// </summary>
    public enum AmountType
    {
        Rate,   // 倍率（例: 0.5 = 基本料金の半額）
        Fixed   // 固定金額（例: 5000円）
    }

    /// <summary>
    /// 金額ありタイプの適用対象料金
    /// </summary>
    public enum TargetFee
    {
        BaseFee,     // 基本料金のみ
        MileageFee,  // 走行距離料金のみ
        Both         // 両方
    }

    /// <summary>
    /// チェックボックスフラグ1件の定義
    /// </summary>
    public class FlagDefinition
    {
        /// <summary>内部識別ID（変更不可）</summary>
        public string Id { get; set; }

        /// <summary>表示名（チェックボックスのラベル）</summary>
        public string DisplayName { get; set; }

        /// <summary>フラグの種類</summary>
        public FlagType Type { get; set; }

        /// <summary>金額ありタイプのみ：計算方法</summary>
        public AmountType? AmountType { get; set; }

        /// <summary>金額ありタイプのみ：値（Rate=倍率、Fixed=円）</summary>
        public double? AmountValue { get; set; }

        /// <summary>金額ありタイプのみ：適用対象料金（BaseFee/MileageFee/Both）</summary>
        public TargetFee TargetFee { get; set; } = TargetFee.BaseFee;

        /// <summary>表示順（1始まり）</summary>
        public int Order { get; set; }

        /// <summary>Excelの列番号（自動割り当て）</summary>
        public int ExcelColumn { get; set; }
    }
}
