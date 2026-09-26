using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HansoInputTool.Services
{
    public class MonthlyRecord
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string ShishaName { get; set; } = "";   // 支社名
        public string VehicleNo  { get; set; } = "";   // 車両番号
        public string VehicleKey   => $"{ShishaName}_{VehicleNo}";
        public string VehicleLabel => $"{ShishaName} {VehicleNo}";
        public double? JitsuzaiSuu { get; set; }        // D列：延実在車輌数
        public double? JitsudouSuu { get; set; }        // E列：延実働車輌数
        public double? Hanso       { get; set; }        // G列：搬送回数
        public double? YuryoKm     { get; set; }        // H列：有料キロ数
        public double? MuryoKm     { get; set; }        // I列：無料キロ数
        public double? Unshu       { get; set; }        // K列：運輸実績
    }

    /// <summary>
    /// チェックリスト表示用の車両エントリ
    /// </summary>
    public class VehicleEntry
    {
        public string Key       { get; set; } = "";   // "{ShishaName}_{VehicleNo}"
        public string Label     { get; set; } = "";   // 表示名
        public string ShishaName{ get; set; } = "";
        public string VehicleNo { get; set; } = "";
        public bool   IsKnown   { get; set; }         // SheetNameMap/FullSheetPatternで解決できた車両
        public bool   IsChecked { get; set; } = true; // チェック状態（デフォルトON）
    }
}
