using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// TransferService の分割定義：集計ファイル(wbShukei)まわり（シート並び替え・月間集計・自動生成）。
    /// 本体・通常シート処理は TransferService.cs / .NormalSheet.cs、給油・東日本シートは .FuelAndEastSheet.cs を参照。
    /// </summary>
    public partial class TransferService
    {
        /// <summary>
        /// [並び替え対応] 集計ファイル(wbShukei)のシート順を支社名ごとに整える。
        /// 「月間集計」は常に先頭、「Template」ひな形シートは常に末尾に固定する。
        /// 廃車済み車両などで自動生成されたシート（末尾に追加される）も、この処理で正しい位置に移動する。
        /// </summary>
        private void ReorderShukeiSheets(ExcelPackage wbShukei)
        {
            try
            {
                var monthly = wbShukei.Workbook.Worksheets.FirstOrDefault(w => w.Name == "月間集計");
                var templateWs = wbShukei.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Template");

                var orderedNames = wbShukei.Workbook.Worksheets
                    .Where(ws => ws.Name != "月間集計" && ws.Name != "Template" && !ws.Name.Contains("登録"))
                    .Select(ws =>
                    {
                        int branchOrder = Array.FindIndex(BranchOrder, b => ws.Name.Contains(b));
                        if (branchOrder < 0) branchOrder = BranchOrder.Length; // 不明な支社名は末尾扱い
                        var m = Regex.Match(ws.Name, @"\d+$");
                        int number = m.Success && int.TryParse(m.Value, out var n) ? n : int.MaxValue;
                        return new { ws.Name, BranchOrder = branchOrder, Number = number };
                    })
                    .OrderBy(v => v.BranchOrder)
                    .ThenBy(v => v.Number)
                    .ThenBy(v => v.Name, StringComparer.Ordinal)
                    .Select(v => v.Name)
                    .ToList();

                // 月間集計を先頭に固定
                if (monthly != null)
                    wbShukei.Workbook.Worksheets.MoveBefore(monthly.Index, 1);

                for (int i = orderedNames.Count - 1; i >= 0; i--)
                {
                    var ws = wbShukei.Workbook.Worksheets.FirstOrDefault(x => x.Name == orderedNames[i]);
                    if (ws == null) continue;
                    if (monthly != null) wbShukei.Workbook.Worksheets.MoveAfter(ws.Index, 1);
                    else                 wbShukei.Workbook.Worksheets.MoveBefore(ws.Index, 1);
                }

                // Templateひな形シートは常に末尾に固定
                if (templateWs != null)
                    wbShukei.Workbook.Worksheets.MoveToEnd(templateWs.Name);

                Logger.Info("集計ファイルのシート順を並べ替えました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "集計ファイルのシート並び替え中にエラーが発生しました。");
            }
        }

        /// <summary>
        /// [月間集計対応] 集計ファイル(wbShukei)の「月間集計」シートを、現在の車両シート構成で作り直す。
        /// ExcelHandler.MonthlySummary.cs の UpdateMonthlySummarySheetIfNeeded と同じロジックだが、
        /// 従来はInput.xlsx側にしか適用されておらずTemplate.xlsx（集計ファイル）側は更新されていなかったため、
        /// 転記の都度こちらでも実行することで、新規追加・廃車済み車両を含めて最新の状態に同期する。
        ///
        /// [デザイン対応] 「月間集計」はExcelの「テーブル」機能（縞模様等の表スタイル）で構成されており、
        /// その表の範囲（ref）は車両数が変わっても自動では広がらない。単純にセルへ値を書き込むだけでは、
        /// 表の範囲外にはみ出した行の見た目（縞模様・罫線・数値の書式）が一切適用されず崩れてしまうため、
        /// EPPlus 8 の ExcelTable.DataRows API でテーブル自体の行数を車両数+合計行に合わせて増減させる。
        /// </summary>
        private void UpdateShukeiMonthlySummary(ExcelPackage wbInput, ExcelPackage wbShukei, FlagDefinitionService flagService, DatabaseService dbService)
        {
            try
            {
                var summarySheet = wbShukei.Workbook.Worksheets.FirstOrDefault(w => w.Name == "月間集計");
                if (summarySheet == null)
                {
                    Logger.Warn("集計ファイルに『月間集計』シートが見つからないため、更新をスキップしました。");
                    return;
                }

                var allVehicleSheets = wbShukei.Workbook.Worksheets
                    .Where(ws => ws.Name != "月間集計" && ws.Name != "Template" && !ws.Name.Contains("登録"))
                    .Select(ws => ws.Name)
                    .Select(name =>
                    {
                        int branchOrder = Array.FindIndex(BranchOrder, b => name.Contains(b));
                        if (branchOrder < 0) branchOrder = BranchOrder.Length;
                        var m = Regex.Match(name, @"\d+$");
                        int number = m.Success && int.TryParse(m.Value, out var n) ? n : int.MaxValue;
                        return new { Name = name, BranchOrder = branchOrder, Number = number };
                    })
                    .OrderBy(v => v.BranchOrder)
                    .ThenBy(v => v.Number)
                    .ThenBy(v => v.Name, StringComparer.Ordinal)
                    .Select(v => v.Name)
                    .ToList();

                var flags = flagService?.Flags ?? new List<FlagDefinition>().AsReadOnly();
                int endCol = 11 + flags.Count;
                // [不要な0対応] 過去に存在したフラグ列など、現在のフラグ数を超えた古いデータが
                // テーブル外の列に残ってしまうことがあるため、少し広めの範囲までクリアする。
                int clearEndCol = Math.Max(endCol, 20);

                // [デザイン対応] 月間集計の表（ヘッダー行の次の行からデータが始まる）の行数を実車両数+合計行に合わせる
                var table = summarySheet.Tables?.FirstOrDefault();
                int dataStartRow;
                if (table != null && allVehicleSheets.Count > 0)
                {
                    int desiredDataRows = allVehicleSheets.Count + 1; // 車両行 + 合計行
                    int currentDataRows = table.Address.End.Row - table.Address.Start.Row; // ヘッダー行を除いた行数

                    // [太字対応] 行を追加する前に、正しい書式を持つ1行目（データの最初の行）のスタイルを控えておく。
                    // ExcelTable.DataRows.AddNewRowsは直前の行（今までの合計行など）の書式をコピーしてしまうため、
                    // 新規追加された行がその書式（太字・罫線）を引き継いでしまう不具合を防ぐ。
                    int normalRowRef = table.Address.Start.Row + 1;
                    var normalStyleIds = new int[clearEndCol];
                    for (int col = 1; col <= clearEndCol; col++)
                        normalStyleIds[col - 1] = summarySheet.Cells[normalRowRef, col].StyleID;

                    if (desiredDataRows > currentDataRows)
                        table.DataRows.AddNewRows(desiredDataRows - currentDataRows);
                    else if (desiredDataRows < currentDataRows)
                        table.DataRows.DeleteRows(desiredDataRows, currentDataRows - desiredDataRows);

                    dataStartRow = table.Address.Start.Row + 1;

                    // 全データ行（合計行になる最終行を含む）に、いったん正しい通常行の書式を適用し直す。
                    // 合計行の太字・上罫線は、この後の書き込み処理で改めて明示的に設定する。
                    for (int row = dataStartRow; row <= table.Address.End.Row; row++)
                        for (int col = 1; col <= clearEndCol; col++)
                            summarySheet.Cells[row, col].StyleID = normalStyleIds[col - 1];
                }
                else
                {
                    // テーブルが見つからない場合は従来通り固定行から開始（見た目の自動調整は行われない）
                    dataStartRow = 4;
                    if (table == null)
                        Logger.Warn("月間集計シートにテーブルが見つからなかったため、表の行数調整はスキップしました。");
                }

                int clearEndRow = table != null ? table.Address.End.Row : dataStartRow + 69;
                for (int row = dataStartRow; row <= clearEndRow; row++)
                    for (int col = 1; col <= clearEndCol; col++)
                    {
                        summarySheet.Cells[row, col].Value = null;
                        summarySheet.Cells[row, col].Formula = null;
                    }

                for (int i = 0; i < allVehicleSheets.Count; i++)
                {
                    string sheetName = allVehicleSheets[i];
                    int currentRow = dataStartRow + i;
                    string branch = BranchOrder.FirstOrDefault(b => sheetName.Contains(b)) ?? "";
                    var numMatch = Regex.Match(sheetName, @"\d+$");
                    object number = numMatch.Success && int.TryParse(numMatch.Value, out int num) ? (object)num : null;

                    try
                    {
                        string safeSheetName = $"'{sheetName}'";

                        summarySheet.Cells[currentRow, 1].Value = $"No.{i + 1}";
                        summarySheet.Cells[currentRow, 2].Value = branch;
                        summarySheet.Cells[currentRow, 3].Value = number;

                        summarySheet.Cells[currentRow, 4].Formula  = $"{safeSheetName}!E4";
                        summarySheet.Cells[currentRow, 5].Formula  = $"{safeSheetName}!G4";
                        summarySheet.Cells[currentRow, 6].Formula  = $"IFERROR(E{currentRow}/D{currentRow},0)";
                        summarySheet.Cells[currentRow, 7].Formula  = $"{safeSheetName}!G4";
                        summarySheet.Cells[currentRow, 8].Formula  = $"{safeSheetName}!H4";
                        summarySheet.Cells[currentRow, 9].Formula  = $"{safeSheetName}!I4";
                        summarySheet.Cells[currentRow, 10].Formula = $"SUM(H{currentRow}:I{currentRow})";
                        summarySheet.Cells[currentRow, 11].Formula = $"{safeSheetName}!K4";

                        // [フラグ情報対応] 実績月報集計にはフラグ情報を記入しない方針になったため、
                        // フラグ列への書き込みは行わない（列自体は上でクリア済み）。
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"月間集計 Row {currentRow} ({sheetName}) のデータ設定中にエラーが発生しました");
                    }
                }

                if (allVehicleSheets.Count > 0)
                {
                    // [デザイン対応] 合計行は必ず最後の車両行の直後（テーブルの最終行）に配置する
                    int totalRow = dataStartRow + allVehicleSheets.Count;
                    int lastDataRow = totalRow - 1;
                    summarySheet.Cells[totalRow, 1].Value = "合　計";
                    summarySheet.Cells[totalRow, 4].Formula  = $"SUM(D{dataStartRow}:D{lastDataRow})";
                    summarySheet.Cells[totalRow, 5].Formula  = $"SUM(E{dataStartRow}:E{lastDataRow})";
                    summarySheet.Cells[totalRow, 6].Formula  = $"IFERROR(E{totalRow}/D{totalRow},0)";
                    summarySheet.Cells[totalRow, 7].Formula  = $"SUM(G{dataStartRow}:G{lastDataRow})";
                    summarySheet.Cells[totalRow, 8].Formula  = $"SUM(H{dataStartRow}:H{lastDataRow})";
                    summarySheet.Cells[totalRow, 9].Formula  = $"SUM(I{dataStartRow}:I{lastDataRow})";
                    summarySheet.Cells[totalRow, 10].Formula = $"SUM(J{dataStartRow}:J{lastDataRow})";
                    summarySheet.Cells[totalRow, 11].Formula = $"SUM(K{dataStartRow}:K{lastDataRow})";

                    // [デザイン対応] 合計行は太字＋上罫線で強調する（通常行と見分けられるように）
                    var totalRange = summarySheet.Cells[totalRow, 1, totalRow, 11];
                    totalRange.Style.Font.Bold = true;
                    totalRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    // [レイアウト対応] 「合計」の文字をA列・B列にまたがって中央に表示する。
                    // ExcelTableの範囲内はセル結合と相性が悪い（行数の自動増減処理と衝突するおそれがある）ため、
                    // セル結合ではなく「選択範囲内で中央」表示＋境界線の非表示で見た目だけ揃える。
                    var totalLabelRange = summarySheet.Cells[totalRow, 1, totalRow, 2];
                    totalLabelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                    summarySheet.Cells[totalRow, 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                    summarySheet.Cells[totalRow, 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                }
                else
                {
                    summarySheet.Cells[dataStartRow, 1].Value = "（車両データなし）";
                }

                wbShukei.Workbook.CalcMode = ExcelCalcMode.Automatic;
                Logger.Info($"集計ファイルの月間集計シートを更新しました（対象車両数: {allVehicleSheets.Count}）。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "集計ファイルの月間集計シート更新中にエラーが発生しました。");
            }
        }

        /// <summary>Excel使用時（DB未使用時）に、Input.xlsx上のフラグ列がONの行数を数える。</summary>
        private int GetFlagCountFromExcel(ExcelPackage wbInput, string sheetName, int excelColumn)
        {
            var ws = wbInput.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
            if (ws == null) return 0;
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return 0;
            int count = 0;
            for (int r = 3; r < totalRowIndex; r++)
                if (GetInt(ws.Cells[r, excelColumn].Value) == 1) count++;
            return count;
        }


        /// <summary>
        /// [廃車済み車両対応] 集計ファイル(wbShukei)に対象シートが存在しない場合に自動生成する。
        /// 廃車等により現在のTemplate.xlsxからは削除済みの車両でも、過去のInput.xlsxを読み込んで
        /// 転記すると出力から漏れてしまっていたため、その場で「Template」シートから複製して対応する。
        /// </summary>
        private ExcelWorksheet EnsureShukeiSheet(ExcelPackage wbShukei, string sheetName)
        {
            // 完全一致 → CH富士吉田の接頭辞なしシート名に対応するEndsWithフォールバックの順に探す
            var existing = wbShukei.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
                        ?? wbShukei.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.EndsWith(sheetName));
            if (existing != null) return existing;

            var templateWs = wbShukei.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Template");
            if (templateWs == null)
            {
                Logger.Warn($"[{sheetName}] は集計ファイルに存在せず、コピー元の『Template』シートも見つからなかったため出力できませんでした。");
                return null;
            }

            var newWs = wbShukei.Workbook.Worksheets.Add(sheetName, templateWs);
            PopulateNewShukeiSheetCells(newWs, sheetName);
            if (!AutoGeneratedVehicleSheets.Contains(sheetName))
                AutoGeneratedVehicleSheets.Add(sheetName);
            Logger.Info($"[{sheetName}] は集計ファイルに存在しなかったため『Template』から自動生成しました（廃車済みなどでTemplate.xlsxに未登録の車両の可能性があります）。");
            return newWs;
        }

        /// <summary>
        /// [廃車済み車両対応] Template.xlsxのSetupNewSheetCells相当の処理。
        /// 新規作成した車両シートの支社名・車輌番号のセルを、シート名から解析して埋める。
        /// 東日本セレモニー系はB4=支社名／C4=車輌番号、それ以外はD1=支社名+種別／H1=車輌番号。
        /// </summary>
        private void PopulateNewShukeiSheetCells(ExcelWorksheet ws, string sheetName)
        {
            // [不具合修正] 実際のシート構成を確認したところ、集計ファイルの車両シートは
            // 車種に関わらずB4(支社名)・C4(車輌番号)のみを使用しており、1行目にD1/H1のような
            // 項目は存在しない（1行目はA1=元号,B1=月,C1='月分'のみ）。誤って余分な文字を
            // 入れてしまっていたため、D1/H1への書き込みは廃止する。
            string branch = BranchOrder.FirstOrDefault(b => sheetName.Contains(b)) ?? sheetName.Trim();
            var numberMatch = Regex.Match(sheetName, @"\d+$");
            int? vehicleNumber = numberMatch.Success && int.TryParse(numberMatch.Value, out int num) ? num : (int?)null;

            ws.Cells["B4"].Value = branch;
            if (vehicleNumber.HasValue)
                ws.Cells["C4"].Value = vehicleNumber.Value;
        }
    }
}
