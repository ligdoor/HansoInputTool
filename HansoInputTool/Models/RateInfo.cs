using Newtonsoft.Json;

namespace HansoInputTool.Models
{
    public class RateInfo
    {
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