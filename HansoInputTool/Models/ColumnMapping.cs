using System.Collections.Generic;
using Newtonsoft.Json;

namespace HansoInputTool.Models
{
    public class ColumnMapping
    {
        public SheetColumnMap NormalSheet { get; set; }
        public CellAddressMap EastSheet   { get; set; }
        public CellAddressMap ShukeiSheet { get; set; }
    }

    public class SheetColumnMap
    {
        public int Day           { get; set; }
        public int HansoCount    { get; set; }
        public int YuryoKm       { get; set; }
        public int MuryoKm       { get; set; }
        public int KihonFee      { get; set; }
        public int SokoFee       { get; set; }
        public int ShinyaFee     { get; set; }
        public int TotalFee      { get; set; }
        public int ShinyaMinutes { get; set; }

        // IsKoryo・IsEmbalmingは後方互換用に残す（FlagDefinitionServiceが主）
        [JsonIgnore] public int IsKoryo    => 12;
        [JsonIgnore] public int IsEmbalming => 13;
    }

    public class CellAddressMap
    {
        // EastSheet
        public string Jitsudo     { get; set; }
        public string Hanso       { get; set; }
        public string UnsoJisseki { get; set; }

        // ShukeiSheet
        public string Days   { get; set; }
        public string YuryoKm { get; set; }
        public string MuryoKm { get; set; }
        public string Total   { get; set; }
    }
}
