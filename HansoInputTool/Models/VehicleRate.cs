using Newtonsoft.Json;

namespace HansoInputTool.Models
{
    // rates.json の新しいデータ構造に対応するクラス
    public class VehicleRate
    {
        [JsonProperty("VehicleTypeName")]
        public string VehicleTypeName { get; set; }

        [JsonProperty("BaseFee")]
        public int BaseFee { get; set; }

        [JsonProperty("MileageFee")]
        public int MileageFee { get; set; }

        [JsonProperty("LateNightFixedFee")]
        public int LateNightFixedFee { get; set; }

        [JsonProperty("LateNightUnitFee")]
        public int LateNightUnitFee { get; set; }
    }
}