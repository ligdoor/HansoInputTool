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
    /// <summary>
    /// VehicleAnnualSummaryService の分割定義：日付範囲・数値変換・シート名解析の小さなヘルパー群。
    /// </summary>
    public partial class VehicleAnnualSummaryService
    {
        // ---- ヘルパー ----

        private static bool IsInRange(int y, int m, int sy, int sm, int ey, int em)
        {
            int val = y * 100 + m;
            return val >= sy * 100 + sm && val <= ey * 100 + em;
        }

        private static List<(int Year, int Month)> GetMonthRange(int sy, int sm, int ey, int em)
        {
            var list = new List<(int, int)>();
            int y = sy, m = sm;
            while (y * 100 + m <= ey * 100 + em)
            {
                list.Add((y, m));
                if (++m > 12) { m = 1; y++; }
            }
            return list;
        }

        private static double? GetDouble(object val)
        {
            if (val == null) return null;
            var s = val.ToString();
            if (s.StartsWith("=")) return null;
            return double.TryParse(s, out double d) ? d : null;
        }

        private static bool TryParseSheetName(string sname, out string shisha, out string vehicleNo)
        {
            shisha = null; vehicleNo = null;
            if (SheetNameMap.TryGetValue(sname, out var mapped))
            {
                shisha = mapped.Shisha; vehicleNo = mapped.VehicleNo; return true;
            }
            var m = FullSheetPattern.Match(sname);
            if (!m.Success) return false;
            shisha = m.Groups[1].Value; vehicleNo = m.Groups[2].Value; return true;
        }
    }
}
