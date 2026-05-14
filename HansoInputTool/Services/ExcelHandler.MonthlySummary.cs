using System;
using System.Linq;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// ExcelHandler の partial クラス：月間集計シートの更新処理
    /// </summary>
    public partial class ExcelHandler
    {
        private void UpdateMonthlySummarySheetIfNeeded(ExcelPackage package)
        {
            var summarySheet = package.Workbook.Worksheets["月間集計"];
            if (summarySheet == null)
            {
                Logger.Warn("月間集計シートが見つかりません。");
                return;
            }

            Logger.Info("=== 月間集計シートの更新を開始 ===");

            var allVehicleSheets = package.Workbook.Worksheets
                .Where(ws => ws.Name != "月間集計")
                .Select(ws => ws.Name)
                .OrderBy(s => GetCategoryOrder(s))
                .ThenBy(s => s)
                .ToList();

            Logger.Info($"対象車両シート数: {allVehicleSheets.Count}");

            const int dataStartRow = 6;
            const int startCol     = 1;  // A列
            const int maxDataRows  = 69;

            // 動的フラグを取得
            var flags  = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();
            int endCol = 11 + flags.Count; // K列まで固定 + フラグ列数分

            // 既存データをクリア
            for (int row = dataStartRow; row < dataStartRow + maxDataRows; row++)
                for (int col = startCol; col <= endCol; col++)
                {
                    summarySheet.Cells[row, col].Value   = null;
                    summarySheet.Cells[row, col].Formula = null;
                }

            Logger.Info($"{maxDataRows}行分のデータをクリアしました");

            // シートごとに行を書き込む
            for (int i = 0; i < allVehicleSheets.Count; i++)
            {
                string sheetName = allVehicleSheets[i];
                var (branch, number) = ParseSheetNameToBranchAndNumber(sheetName);
                int currentRow = dataStartRow + i;

                try
                {
                    string safeSheetName = NeedQuotes(sheetName) ? $"'{sheetName}'" : sheetName;

                    summarySheet.Cells[currentRow, 1].Value = $"No.{i + 1}";
                    summarySheet.Cells[currentRow, 2].Value = branch;
                    summarySheet.Cells[currentRow, 3].Value = int.TryParse(number, out int num) ? num : (object)number;

                    summarySheet.Cells[currentRow, 4].Formula  = $"{safeSheetName}!E4";
                    summarySheet.Cells[currentRow, 5].Formula  = $"{safeSheetName}!G4";
                    summarySheet.Cells[currentRow, 6].Formula  = $"IF(E{currentRow}>0,D{currentRow}/E{currentRow},0)";
                    summarySheet.Cells[currentRow, 7].Formula  = $"{safeSheetName}!G4";
                    summarySheet.Cells[currentRow, 8].Formula  = $"{safeSheetName}!H4";
                    summarySheet.Cells[currentRow, 9].Formula  = $"{safeSheetName}!I4";
                    summarySheet.Cells[currentRow, 10].Formula = $"H{currentRow}+I{currentRow}";
                    summarySheet.Cells[currentRow, 11].Formula = $"{safeSheetName}!K4";

                    // 動的フラグ列（12列目以降）
                    for (int fi = 0; fi < flags.Count; fi++)
                    {
                        int summaryCol = 12 + fi;
                        summarySheet.Cells[currentRow, summaryCol].Value =
                            GetFlagCount(sheetName, flags[fi].ExcelColumn);
                    }

                    Logger.Info($"Row {currentRow}: {sheetName} のデータを設定しました");
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"Row {currentRow} ({sheetName}) のデータ設定中にエラーが発生しました");
                }
            }

            if (!allVehicleSheets.Any())
                summarySheet.Cells[dataStartRow, 1].Value = "（車両データなし）";

            package.Workbook.CalcMode = ExcelCalcMode.Automatic;
            Logger.Info("=== 月間集計シートの更新が完了しました ===");
        }
    }
}
