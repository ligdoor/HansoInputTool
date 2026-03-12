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
        public string ShishaName { get; set; } = "";   // B列：支社名
        public string VehicleNo { get; set; } = "";   // C列：車両番号
        public string VehicleKey => $"{ShishaName}_{VehicleNo}";
        public string VehicleLabel => $"{ShishaName} {VehicleNo}";
        public double? Unshu { get; set; }             // K列：運輸実績
    }

    public class VehicleAnnualSummaryService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private const string TargetSheetName = "月間集計";
        private const int DataStartRow = 4;   // 4行目からデータ開始
        private const int ColShisha = 2;   // B列：支社名
        private const int ColVehicle = 3;   // C列：車両番号
        private const int ColUnshu = 11;  // K列：運輸実績

        // ファイル名パターン: "#期 #月 R# アルス搬送・霊柩車　実績月報集計.xlsx"
        // 例: "44期 5月 R6 アルス搬送・霊柩車　実績月報集計.xlsx"
        private static readonly Regex FilePattern = new Regex(
            @"\d+期\s+(\d+)月\s+R(\d+)\s+アルス搬送・霊柩車[\s　]+実績月報集計\.xlsx$",
            RegexOptions.Compiled);

        /// <summary>
        /// 最上位フォルダ配下を再帰検索して対象期間のデータを読み込む
        /// </summary>
        public List<MonthlyRecord> LoadData(
            string rootFolder,
            int startYear, int startMonth,
            int endYear, int endMonth)
        {
            var result = new List<MonthlyRecord>();

            if (!Directory.Exists(rootFolder))
            {
                Logger.Warn($"フォルダが存在しません: {rootFolder}");
                return result;
            }

            // 全サブフォルダを再帰的に検索
            var files = Directory.GetFiles(rootFolder, "*実績月報集計*.xlsx",
                                           SearchOption.AllDirectories);

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var match = FilePattern.Match(fileName);
                if (!match.Success) continue;

                int month = int.Parse(match.Groups[1].Value);  // 第1グループ：月
                int reiwa = int.Parse(match.Groups[2].Value);  // 第2グループ：令和年
                int year = 2018 + reiwa;  // 令和→西暦

                if (!IsInRange(year, month, startYear, startMonth, endYear, endMonth))
                    continue;

                Logger.Info($"読み込み: {filePath} ({year}年{month}月)");

                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using var pkg = new ExcelPackage(new FileInfo(filePath));

                    var ws = pkg.Workbook.Worksheets
                        .FirstOrDefault(s => s.Name == TargetSheetName);
                    if (ws == null)
                    {
                        Logger.Warn($"「{TargetSheetName}」シートなし: {fileName}");
                        continue;
                    }

                    for (int row = DataStartRow; ; row++)
                    {
                        var shisha = ws.Cells[row, ColShisha].Value?.ToString()?.Trim() ?? "";

                        // 空行で終了
                        if (string.IsNullOrEmpty(shisha)) break;

                        // 「合計」行はスキップ
                        if (shisha.Contains("合計") || shisha.Contains("合　計")) continue;

                        var vehicleNo = ws.Cells[row, ColVehicle].Value?.ToString()?.Trim() ?? "";
                        var unshu = GetDouble(ws.Cells[row, ColUnshu].Value);

                        result.Add(new MonthlyRecord
                        {
                            Year = year,
                            Month = month,
                            ShishaName = shisha,
                            VehicleNo = vehicleNo,
                            Unshu = unshu,
                        });
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"ファイル読み込みエラー: {filePath}");
                }
            }

            return result;
        }

        /// <summary>
        /// 集計結果をExcelに出力する（縦=月、横=車両）
        /// </summary>
        public void ExportToExcel(
            List<MonthlyRecord> allData,
            string outputPath,
            int startYear, int startMonth,
            int endYear, int endMonth)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pkg = new ExcelPackage();
            var ws = pkg.Workbook.Worksheets.Add("車両別年度集計");

            var months = GetMonthRange(startYear, startMonth, endYear, endMonth);

            // 車両リスト：支社名→車両番号の順でソート
            var vehicles = allData
                .Select(d => (d.VehicleKey, d.VehicleLabel, d.ShishaName, d.VehicleNo))
                .Distinct()
                .OrderBy(v => v.ShishaName)
                .ThenBy(v => v.VehicleNo)
                .ToList();

            int totalCol = vehicles.Count + 2;  // 合計列の位置

            // ====== ヘッダー行 ======

            // 1行目：タイトル
            ws.Cells[1, 1].Value =
                $"運輸実績　{startYear}年{startMonth}月 〜 {endYear}年{endMonth}月";
            ws.Cells[1, 1, 1, totalCol].Merge = true;
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 13;
            ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // 2行目：列ヘッダー
            ws.Cells[2, 1].Value = "年月";
            SetHeaderStyle(ws.Cells[2, 1]);

            for (int vi = 0; vi < vehicles.Count; vi++)
            {
                ws.Cells[2, vi + 2].Value = vehicles[vi].VehicleLabel;
                SetHeaderStyle(ws.Cells[2, vi + 2]);
            }

            ws.Cells[2, totalCol].Value = "合　計";
            SetHeaderStyle(ws.Cells[2, totalCol]);

            // ====== データ行（月ごと）======

            for (int mi = 0; mi < months.Count; mi++)
            {
                int dataRow = mi + 3;
                var (y, m) = months[mi];

                // 年月ラベル
                ws.Cells[dataRow, 1].Value = $"{y}年{m}月";
                ws.Cells[dataRow, 1].Style.Font.Bold = true;
                ws.Cells[dataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // 各車両の値
                for (int vi = 0; vi < vehicles.Count; vi++)
                {
                    var record = allData.FirstOrDefault(d =>
                        d.VehicleKey == vehicles[vi].VehicleKey
                        && d.Year == y && d.Month == m);

                    if (record?.Unshu is double val && val != 0)
                    {
                        ws.Cells[dataRow, vi + 2].Value = val;
                        ws.Cells[dataRow, vi + 2].Style.Numberformat.Format = "#,##0";
                    }
                }

                // 行合計（SUM式）
                string rangeAddr = $"{ws.Cells[dataRow, 2].Address}:{ws.Cells[dataRow, totalCol - 1].Address}";
                ws.Cells[dataRow, totalCol].Formula = $"SUM({rangeAddr})";
                ws.Cells[dataRow, totalCol].Style.Numberformat.Format = "#,##0";
                ws.Cells[dataRow, totalCol].Style.Font.Bold = true;
            }

            // ====== 合計行 ======

            int totalRow = months.Count + 3;
            int firstData = 3;
            int lastData = months.Count + 2;

            ws.Cells[totalRow, 1].Value = "合　計";
            ws.Cells[totalRow, 1].Style.Font.Bold = true;
            ws.Cells[totalRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            for (int vi = 0; vi < vehicles.Count; vi++)
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

            // 合計行の背景色（薄い青）
            var totalRange = ws.Cells[totalRow, 1, totalRow, totalCol];
            totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalRange.Style.Fill.BackgroundColor.SetColor(
                System.Drawing.Color.FromArgb(219, 234, 254));

            // ====== 書式仕上げ ======

            // 全体に薄い罫線
            var dataRange = ws.Cells[2, 1, totalRow, totalCol];
            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

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
            return double.TryParse(val.ToString(), out double d) ? d : null;
        }

        private static void SetHeaderStyle(ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(
                System.Drawing.Color.FromArgb(30, 58, 138));  // 紺
            cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }
    }
}