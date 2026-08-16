using System.Collections.Generic;
using System.Linq;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.Models
{
    public class RowData : ObservableObject
    {
        public int RowIndex { get; set; }

        /// <summary>SQLite の主キー（DB使用時のみ設定。Excel使用時は 0）</summary>
        public long DbId { get; set; }
        public int? B_Day { get; set; }
        public int? C_Hanso { get; set; }
        public int? D_YuryoKm { get; set; }
        public int? E_MuryoKm { get; set; }
        public int? H_LateFeeOotsuki { get; set; }
        public int? K_LateMinutes { get; set; }
        public string LateValueText { get; set; }

        // 動的フラグ: FlagDefinition.Id → 1(ON) or null
        public Dictionary<string, int?> FlagValues { get; set; } = new();

        // フラグがONかどうかを Id で取得するヘルパー
        public bool GetFlag(string flagId)
            => FlagValues.TryGetValue(flagId, out var v) && v == 1;

        // プレビュー用テキスト（ONのフラグ表示名をカンマ区切りで返す）
        // FlagDefinitionsはプレビュー更新時に外部からセットする
        public IReadOnlyList<FlagDefinition> FlagDefinitions { get; set; }

        public string FlagSummaryText
        {
            get
            {
                if (FlagDefinitions == null || FlagValues == null) return string.Empty;
                var active = FlagDefinitions
                    .Where(f => FlagValues.TryGetValue(f.Id, out var v) && v == 1)
                    .Select(f => f.DisplayName);
                return string.Join(", ", active);
            }
        }

        // 後方互換用（既存XAMLバインディングが残っている場合のため）
        public int? L_IsKoryo    => FlagValues.TryGetValue("koryo",     out var k) ? k : null;
        public int? M_IsEmbalming => FlagValues.TryGetValue("embalming", out var e) ? e : null;
        public string IsKoryoText     => L_IsKoryo    == 1 ? "✔" : "";
        public string IsEmbalmingText => M_IsEmbalming == 1 ? "✔" : "";

        /// <summary>この日の給油記録の要約テキスト（例: "⛽12,345km/40L"）。給油記録が無ければ空文字。</summary>
        public string FuelSummaryText { get; set; } = string.Empty;
    }
}
