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
    /// VehicleAnnualSummaryService の分割定義：Excel出力(ExportToExcel)とシート書式まわり。
    /// </summary>
    public partial class VehicleAnnualSummaryService
    {
        // ─────────────────────────────────────
        // Step3: 選択車両でExcel出力
        // annual_results.xlsx（ひな形）を実際に開き、その中で「Template」シートを
        // 月ごとに複製して「月間集計 R#.#月」を作成、最後に「年間実績」シートへ
        // 全期間を3D集計する。デザイン（テーブル・配色・書式）はひな形のものをそのまま使う。
        // ─────────────────────────────────────
        public void ExportToExcel(
            List<MonthlyRecord> allData,
            List<VehicleEntry>  selectedVehicles,
            string outputPath,
            int startYear, int startMonth,
            int endYear,   int endMonth,
            string annualTemplatePath)
        {
            if (string.IsNullOrWhiteSpace(annualTemplatePath) || !File.Exists(annualTemplatePath))
                throw new FileNotFoundException($"年間集計のひな形ファイルが見つかりません: {annualTemplatePath}");

            // [EPPlus 8対応] LicenseContextはEPPlus 8で非推奨（obsolete）になったため、
            // 新しいLicense APIに変更。※「アルス」の部分は実際の組織名に置き換えてください。
            ExcelPackage.License.SetNonCommercialOrganization("アルス");

            // [デザイン維持対応] ひな形ファイル自体を開いて作業する。EPPlusには「テーブルを含む
            // シートを別ファイルへコピーするとファイルが壊れる」既知の不具合があるため、
            // 新規シートの複製はすべて「同じファイル（同じパッケージ）内」で行い、
            // 最後に別名（outputPath）で保存する（ひな形ファイル自体は一切上書きしない）。
            using var pkg = new ExcelPackage(new FileInfo(annualTemplatePath));

            var templateWs = pkg.Workbook.Worksheets["Template"];
            var summaryWs  = pkg.Workbook.Worksheets["年間実績"];
            if (templateWs == null || summaryWs == null)
                throw new InvalidOperationException(
                    "ひな形ファイル(annual_results.xlsx)に「Template」または「年間実績」シートが見つかりません。");

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var months = GetMonthRange(startYear, startMonth, endYear, endMonth);
            int vehicleCount = selectedVehicles.Count;
            int totalDataRow = DataStartRow + vehicleCount; // 合計行

            var monthlySheetNames = new List<string>();

            // ---- 月ごとに「Template」シートを複製 ----
            foreach (var (y, m) in months)
            {
                int eraNum = y - (eraStartYear - 1);
                string sheetName = $"月間集計 {eraName}{eraNum}.{m}月";
                monthlySheetNames.Add(sheetName);

                var ws = pkg.Workbook.Worksheets.Add(sheetName, templateWs);
                ResizeSheetTable(ws, vehicleCount);

                ws.Cells[1, 1].Value = $"{eraName}{eraNum}";
                ws.Cells[1, 2].Value = m; // B1：月の数値
                ws.Cells[1, 3].Value = "月分";

                for (int vi = 0; vi < vehicleCount; vi++)
                {
                    var v = selectedVehicles[vi];
                    int row = DataStartRow + vi;
                    var record = allData.FirstOrDefault(d =>
                        d.VehicleKey == v.Key && d.Year == y && d.Month == m);

                    WriteVehicleRowValues(ws, row, vi + 1, v.ShishaName, v.VehicleNo, record);
                }

                WriteTotalRowFormulas(ws, totalDataRow, DataStartRow, totalDataRow - 1);
            }

            // ---- 年間実績シート（ひな形に元から入っているシートを流用） ----
            ResizeSheetTable(summaryWs, vehicleCount);

            int startEraNum = startYear - (eraStartYear - 1);
            int endEraNum   = endYear   - (eraStartYear - 1);
            summaryWs.Cells[1, 1].Value =
                $"{eraName}{startEraNum}年{startMonth}月～{eraName}{endEraNum}年{endMonth}月";

            string sheetRange3D = monthlySheetNames.Count == 1
                ? $"'{monthlySheetNames[0]}'"
                : $"'{monthlySheetNames[0]}:{monthlySheetNames[^1]}'";

            for (int vi = 0; vi < vehicleCount; vi++)
            {
                var v = selectedVehicles[vi];
                int row = DataStartRow + vi;

                summaryWs.Cells[row, 1].Value = $"№{vi + 1}";
                summaryWs.Cells[row, 2].Value = v.ShishaName;
                summaryWs.Cells[row, 3].Value = int.TryParse(v.VehicleNo, out int vn) ? (object)vn : v.VehicleNo;

                summaryWs.Cells[row, 4].Formula  = $"SUM({sheetRange3D}!D{row})";
                summaryWs.Cells[row, 5].Formula  = $"SUM({sheetRange3D}!E{row})";
                summaryWs.Cells[row, 6].Formula  = $"IFERROR(E{row}/(D{row}*1),0)";
                summaryWs.Cells[row, 7].Formula  = $"SUM({sheetRange3D}!G{row})";
                summaryWs.Cells[row, 8].Formula  = $"SUM({sheetRange3D}!H{row})";
                summaryWs.Cells[row, 9].Formula  = $"SUM({sheetRange3D}!I{row})";
                summaryWs.Cells[row, 10].Formula = $"SUM(H{row},I{row})";
                summaryWs.Cells[row, 11].Formula = $"SUM({sheetRange3D}!K{row})";
            }

            WriteTotalRowFormulas(summaryWs, totalDataRow, DataStartRow, totalDataRow - 1);

            // ひな形として使った「Template」シートは出力には不要なので削除し、
            // 「年間実績」はタブの一番最後に配置する（複製した月次シートは複製順のまま先頭側に並ぶ）。
            pkg.Workbook.Worksheets.Delete(templateWs);
            pkg.Workbook.Worksheets.MoveToEnd("年間実績");

            pkg.SaveAs(new FileInfo(outputPath));
            Logger.Info($"年間実績Excel出力完了: {outputPath}（月次シート{monthlySheetNames.Count}枚＋年間実績、車両{vehicleCount}台）");
        }

        /// <summary>
        /// [デザイン維持対応] シート内のExcelテーブルの行数を、車両数+合計行に合わせて増減させる。
        /// AddNewRows/DeleteRowsは直前の行の書式をそのまま新しい行にコピーしてしまうことがあるため、
        /// リサイズ前に「通常データ行」と「元の合計行」双方の書式を控えておき、リサイズ後に
        /// 該当する行へ正しい書式を明示的に再適用する（縞模様自体はテーブル機能側で自動描画される）。
        /// </summary>
        private static void ResizeSheetTable(ExcelWorksheet ws, int vehicleCount)
        {
            var table = ws.Tables?.FirstOrDefault();
            if (table == null) return;

            int desiredDataRows = vehicleCount + 1; // 車両行 + 合計行
            int currentDataRows = table.Address.End.Row - table.Address.Start.Row;
            int firstCol = table.Address.Start.Column;
            int lastCol  = table.Address.End.Column;
            int normalRowRef   = table.Address.Start.Row + 1; // 既存の先頭データ行（通常行の書式見本）
            int oldTotalRowRef = table.Address.End.Row;       // 既存の最終行（合計行の書式見本）

            int width = lastCol - firstCol + 1;
            var normalStyleIds = new int[width];
            var totalStyleIds  = new int[width];
            for (int col = firstCol; col <= lastCol; col++)
            {
                normalStyleIds[col - firstCol] = ws.Cells[normalRowRef, col].StyleID;
                totalStyleIds[col - firstCol]  = ws.Cells[oldTotalRowRef, col].StyleID;
            }

            if (desiredDataRows > currentDataRows)
                table.DataRows.AddNewRows(desiredDataRows - currentDataRows);
            else if (desiredDataRows < currentDataRows)
                table.DataRows.DeleteRows(desiredDataRows, currentDataRows - desiredDataRows);

            int newTotalRow = table.Address.End.Row;
            for (int row = table.Address.Start.Row + 1; row <= table.Address.End.Row; row++)
            {
                var styleSet = (row == newTotalRow) ? totalStyleIds : normalStyleIds;
                for (int col = firstCol; col <= lastCol; col++)
                    ws.Cells[row, col].StyleID = styleSet[col - firstCol];
            }
        }

        /// <summary>
        /// 車両1台分のデータ行に値・数式を書き込む（書式はResizeSheetTableで既に整えられている前提）。
        /// D列（延実在車輌数）はひな形と同じくA1/B1から月の日数を計算する数式のまま、
        /// J列（総走行ｋｍ）はテーブル名参照の代わりに単純な範囲参照に置き換えている
        /// （シート複製後もテーブル名に依存せず安全に計算されるようにするため。表示・計算結果は同じ）。
        /// </summary>
        private static void WriteVehicleRowValues(ExcelWorksheet ws, int row, int no, string shisha, string vehicleNo, MonthlyRecord record)
        {
            ws.Cells[row, 1].Value = $"№{no}";
            ws.Cells[row, 2].Value = shisha;
            ws.Cells[row, 3].Value = int.TryParse(vehicleNo, out int vn) ? (object)vn : vehicleNo;
            ws.Cells[row, 4].Formula = $"DAY(EOMONTH(DATE(VALUE(MID($A$1,2,LEN($A$1)-1))+2018,$B$1,1),0))";
            ws.Cells[row, 5].Value   = record?.JitsudouSuu ?? 0;
            ws.Cells[row, 6].Formula = $"IFERROR(E{row}/(D{row}*1),0)";
            ws.Cells[row, 7].Value   = record?.Hanso ?? 0;
            ws.Cells[row, 8].Value   = record?.YuryoKm ?? 0;
            ws.Cells[row, 9].Value   = record?.MuryoKm ?? 0;
            ws.Cells[row, 10].Formula = $"SUM(H{row}:I{row})";
            ws.Cells[row, 11].Value  = record?.Unshu ?? 0;
        }

        /// <summary>合計行の数式を書き込む（書式はResizeSheetTableで既に整えられている前提）。</summary>
        private static void WriteTotalRowFormulas(ExcelWorksheet ws, int totalRow, int firstDataRow, int lastDataRow)
        {
            ws.Cells[totalRow, 1].Value = "合\u3000\u3000\u3000計";
            ws.Cells[totalRow, 4].Formula  = $"SUM(D{firstDataRow}:D{lastDataRow})";
            ws.Cells[totalRow, 5].Formula  = $"SUM(E{firstDataRow}:E{lastDataRow})";
            ws.Cells[totalRow, 6].Formula  = $"IFERROR(E{totalRow}/(D{totalRow}*1),0)";
            ws.Cells[totalRow, 7].Formula  = $"SUM(G{firstDataRow}:G{lastDataRow})";
            ws.Cells[totalRow, 8].Formula  = $"SUM(H{firstDataRow}:H{lastDataRow})";
            ws.Cells[totalRow, 9].Formula  = $"SUM(I{firstDataRow}:I{lastDataRow})";
            ws.Cells[totalRow, 10].Formula = $"SUM(J{firstDataRow}:J{lastDataRow})";
            ws.Cells[totalRow, 11].Formula = $"SUM(K{firstDataRow}:K{lastDataRow})";
        }

    }
}
