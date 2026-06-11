using System.Collections.Generic;

namespace HansoInputTool.Models
{
    /// <summary>
    /// 車両ごとの個別設定
    /// </summary>
    public class VehicleConfig
    {
        /// <summary>
        /// 深夜入力方式: "time"=深夜時間（分）、"fee"=深夜料金（円）
        /// </summary>
        public string LateInputMode { get; set; } = "time";
    }

    /// <summary>
    /// vehicle_settings.json のルート
    /// キー: シート名（例: "CH大月 寝台車 1603"）
    /// </summary>
    public class VehicleSettings : Dictionary<string, VehicleConfig>
    {
        public bool IsFeeMode(string sheetName)
        {
            if (TryGetValue(sheetName, out var cfg))
                return cfg.LateInputMode == "fee";
            return false;
        }
    }
}
