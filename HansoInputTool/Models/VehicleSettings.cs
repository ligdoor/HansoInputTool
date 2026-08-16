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

        /// <summary>この車両が給油管理表への記録対象かどうか（例: CH富士吉田の車両）</summary>
        public bool IsFuelTracked { get; set; } = false;
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

        /// <summary>指定シートが給油管理の対象かどうかを返す</summary>
        public bool IsFuelTracked(string sheetName)
        {
            if (TryGetValue(sheetName, out var cfg))
                return cfg.IsFuelTracked;
            return false;
        }
    }
}
