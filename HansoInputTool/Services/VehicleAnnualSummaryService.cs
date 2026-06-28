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
        public double? Unshu { get; set; }             // K列：運輸実績
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

    public class VehicleAnnualSummaryService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private const string TargetSheetName = "月間集計";
        private const int DataStartRow = 4;
        private const int ColUnshu = 11;

        private static readonly Regex FilePattern = new Regex(
            @"\d+期\s+(\d+)月\s+([A-Za-z]{1,3})(\d+)\s+アルス搬送・霊柩車\u3000実績月報集計\.xlsx$",
            RegexOptions.Compiled);

        // 短縮シート名 → (支社名, 車番) マッピング
        private static readonly Dictionary<string, (string Shisha, string VehicleNo)> SheetNameMap =
            new Dictionary<string, (string, string)>
            {
                ["寝台車 29"]        = ("CH富士吉田", "29"),
                ["寝台車 30"]        = ("CH富士吉田", "30"),
                ["霊柩車 40"]        = ("CH富士吉田", "40"),
                ["霊柩車 223"]       = ("CH富士吉田", "223"),
                ["大月 寝台車 1603"] = ("CH大月", "1603"),
                ["大月 霊柩車 2577"] = ("CH大月", "2577"),
                ["東日本セレモニー 2"] = ("東日本セレモニー", "2"),
            };

        private static readonly Regex FullSheetPattern = new Regex(
            @"^(CH富士吉田|CH大月|CH東富士|東日本セレモニー)(?:\s+(?:寝台車|霊柩車))?\s+(\d+)$");

        private static readonly Dictionary<string, int> CategoryOrder =
            new Dictionary<string, int>
            {
                ["CH富士吉田"]       = 1,
                ["CH大月"]           = 2,
                ["CH東富士"]         = 3,
                ["東日本セレモニー"] = 4,
            };

        // ─────────────────────────────────────
        // Step1: フォルダをスキャンして車両リストを収集する
        // （チェックリスト表示用。未知の車両もUnknownとして含める）
        // ─────────────────────────────────────
        public List<VehicleEntry> ScanVehicles(
            string rootFolder,
            int startYear, int startMonth,
            int endYear,   int endMonth)
        {
            var entries = new Dictionary<string, VehicleEntry>(); // key→entry

            if (!Directory.Exists(rootFolder)) return new List<VehicleEntry>();

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var files = Directory.GetFiles(rootFolder, "*実績月報集計*.xlsx",
                                           SearchOption.AllDirectories);

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var match = FilePattern.Match(fileName);
                if (!match.Success) continue;

                int month      = int.Parse(match.Groups[1].Value);
                string fileEra = match.Groups[2].Value;
                int eraNum     = int.Parse(match.Groups[3].Value);

                int baseYear = fileEra.Equals(eraName, StringComparison.OrdinalIgnoreCase)
                    ? eraStartYear - 1 : 2018;
                int year = baseYear + eraNum;

                if (!IsInRange(year, month, startYear, startMonth, endYear, endMonth))
                    continue;

                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using var pkg = new ExcelPackage(new FileInfo(filePath));

                    foreach (var ws in pkg.Workbook.Worksheets)
                    {
                        var sname = ws.Name;
                        if (sname == TargetSheetName || sname.Contains("登録") ||
                            sname == "Template" || sname == "月間集計") continue;

                        bool isKnown = TryParseSheetName(sname, out var shisha, out var vehicleNo);

                        if (!isKnown)
                        {
                            // 未知シートも「未分類」として取り込む
                            shisha    = "未分類";
                            vehicleNo = sname;
                            Logger.Warn($"未知のシート名（未分類として追加）: [{sname}]");
                        }

                        var key = $"{shisha}_{vehicleNo}";
                        if (!entries.ContainsKey(key))
                        {
                            entries[key] = new VehicleEntry
                            {
                                Key        = key,
                                Label      = isKnown ? $"{shisha} {vehicleNo}" : $"[未分類] {sname}",
                                ShishaName = shisha,
                                VehicleNo  = vehicleNo,
                                IsKnown    = isKnown,
                                IsChecked  = true,
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"車両スキャンエラー: {filePath}");
                }
            }

            // ソート: 既知→未分類、支社順、車番順
            return entries.Values
                .OrderBy(e => e.IsKnown ? 0 : 1)
                .ThenBy(e => CategoryOrder.TryGetValue(e.ShishaName, out int c) ? c : 99)
                .ThenBy(e => int.TryParse(e.VehicleNo, out int n) ? n : 0)
                .ThenBy(e => e.VehicleNo)
                .ToList();
        }

        // ─────────────────────────────────────
        // Step2: チェックされた車両のみデータを読み込む
        // ─────────────────────────────────────
        public List<MonthlyRecord> LoadData(
            string rootFolder,
            int startYear, int startMonth,
            int endYear,   int endMonth,
            IEnumerable<VehicleEntry> selectedVehicles)
        {
            var result = new List<MonthlyRecord>();
            var selectedKeys = new HashSet<string>(selectedVehicles.Select(v => v.Key));
            if (selectedKeys.Count == 0) return result;

            if (!Directory.Exists(rootFolder)) return result;

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var files = Directory.GetFiles(rootFolder, "*実績月報集計*.xlsx",
                                           SearchOption.AllDirectories);

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var match = FilePattern.Match(fileName);
                if (!match.Success) continue;

                int month      = int.Parse(match.Groups[1].Value);
                string fileEra = match.Groups[2].Value;
                int eraNum     = int.Parse(match.Groups[3].Value);

                int baseYear = fileEra.Equals(eraName, StringComparison.OrdinalIgnoreCase)
                    ? eraStartYear - 1 : 2018;
                int year = baseYear + eraNum;

                if (!IsInRange(year, month, startYear, startMonth, endYear, endMonth))
                    continue;

                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using var pkg = new ExcelPackage(new FileInfo(filePath));

                    foreach (var ws in pkg.Workbook.Worksheets)
                    {
                        var sname = ws.Name;
                        if (sname == TargetSheetName || sname.Contains("登録") ||
                            sname == "Template" || sname == "月間集計") continue;

                        bool isKnown = TryParseSheetName(sname, out var shisha, out var vehicleNo);
                        if (!isKnown) { shisha = "未分類"; vehicleNo = sname; }

                        var key = $"{shisha}_{vehicleNo}";
                        if (!selectedKeys.Contains(key)) continue; // チェックOFFはスキップ

                        var unshu = GetDouble(ws.Cells[DataStartRow, ColUnshu].Value);
                        if (unshu == null) continue;

                        result.Add(new MonthlyRecord
                        {
                            Year       = year,
                            Month      = month,
                            ShishaName = shisha,
                            VehicleNo  = vehicleNo,
                            Unshu      = unshu,
                        });
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"データ読み込みエラー: {filePath}");
                }
            }

            return result;
        }

        // ─────────────────────────────────────
        // Step3: 選択車両でExcel出力
        // ─────────────────────────────────────
        public void ExportToExcel(
            List<MonthlyRecord> allData,
            List<VehicleEntry>  selectedVehicles,
            string outputPath,
            int startYear, int startMonth,
            int endYear,   int endMonth)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pkg = new ExcelPackage();
            var ws = pkg.Workbook.Worksheets.Add("車両別年度集計");

            var months = GetMonthRange(startYear, startMonth, endYear, endMonth);
            int totalCol = selectedVehicles.Count + 2;

            // ヘッダー1行目
            ws.Cells[1, 1].Value =
                $"運輸実績　{startYear}年{startMonth}月 〜 {endYear}年{endMonth}月";
            ws.Cells[1, 1, 1, totalCol].Merge = true;
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 13;
            ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // ヘッダー2行目
            ws.Cells[2, 1].Value = "年月";
            SetHeaderStyle(ws.Cells[2, 1]);
            for (int vi = 0; vi < selectedVehicles.Count; vi++)
            {
                ws.Cells[2, vi + 2].Value = selectedVehicles[vi].Label;
                SetHeaderStyle(ws.Cells[2, vi + 2]);
            }
            ws.Cells[2, totalCol].Value = "合　計";
            SetHeaderStyle(ws.Cells[2, totalCol]);

            // データ行
            for (int mi = 0; mi < months.Count; mi++)
            {
                int dataRow = mi + 3;
                var (y, m) = months[mi];

                ws.Cells[dataRow, 1].Value = $"{y}年{m}月";
                ws.Cells[dataRow, 1].Style.Font.Bold = true;
                ws.Cells[dataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                for (int vi = 0; vi < selectedVehicles.Count; vi++)
                {
                    var record = allData.FirstOrDefault(d =>
                        d.VehicleKey == selectedVehicles[vi].Key
                        && d.Year == y && d.Month == m);
                    if (record?.Unshu is double val && val != 0)
                    {
                        ws.Cells[dataRow, vi + 2].Value = val;
                        ws.Cells[dataRow, vi + 2].Style.Numberformat.Format = "#,##0";
                    }
                }

                string rangeAddr = $"{ws.Cells[dataRow, 2].Address}:{ws.Cells[dataRow, totalCol - 1].Address}";
                ws.Cells[dataRow, totalCol].Formula = $"SUM({rangeAddr})";
                ws.Cells[dataRow, totalCol].Style.Numberformat.Format = "#,##0";
                ws.Cells[dataRow, totalCol].Style.Font.Bold = true;
            }

            // 合計行
            int totalRow  = months.Count + 3;
            int firstData = 3;
            int lastData  = months.Count + 2;

            ws.Cells[totalRow, 1].Value = "合　計";
            ws.Cells[totalRow, 1].Style.Font.Bold = true;
            ws.Cells[totalRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int vi = 0; vi < selectedVehicles.Count; vi++)
            {
                int col = vi + 2;
                ws.Cells[totalRow, col].Formula =
                    $"SUM({ws.Cells[firstData, col].Address}:{ws.Cells[lastData, col].Address})";
                ws.Cells[totalRow, col].Style.Numberformat.Format = "#,##0";
                ws.Cells[totalRow, col].Style.Font.Bold = true;
            }
            ws.Cells[totalRow, totalCol].Formula =
                $"SUM({ws.Cells[firstData, totalCol].Address}:{ws.Cells[lastData, totalCol].Address})";
            ws.Cells[totalRow, totalCol].Style.Numberformat.Format = "#,##0";
            ws.Cells[totalRow, totalCol].Style.Font.Bold = true;

            var totalRange = ws.Cells[totalRow, 1, totalRow, totalCol];
            totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalRange.Style.Fill.BackgroundColor.SetColor(
                System.Drawing.Color.FromArgb(219, 234, 254));

            var dataRange = ws.Cells[2, 1, totalRow, totalCol];
            dataRange.Style.Border.Top.Style    = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style   = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style  = ExcelBorderStyle.Thin;

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            ws.Column(1).Width = Math.Max(ws.Column(1).Width, 12);

            pkg.SaveAs(new FileInfo(outputPath));
            Logger.Info($"集計Excel出力完了: {outputPath}");
        }

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

        private static void SetHeaderStyle(ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(
                System.Drawing.Color.FromArgb(30, 58, 138));
            cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }
    }
}
