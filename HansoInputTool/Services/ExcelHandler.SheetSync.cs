using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// ExcelHandler の partial クラス：シート同期・並べ替え・命名処理
    /// </summary>
    public partial class ExcelHandler
    {
        #region 公開API

        public void SyncAllVehicleSheets(
            List<string> sheetsToDelete,
            Dictionary<string, string> renameMap,
            List<(string newName, string templateName)> sheetsToAdd)
        {
            SyncPackageSheets(_inputPackage,    "Input.xlsx",    sheetsToDelete, renameMap, sheetsToAdd, isInputFile: true);
            SyncPackageSheets(_templatePackage, "Template.xlsx", sheetsToDelete, renameMap, sheetsToAdd, isInputFile: false);

            ReorderVehicleSheets(_inputPackage);
            ReorderVehicleSheets(_templatePackage);

            // 月間集計は Input 側のみ更新（Template を破壊しない）
            UpdateMonthlySummarySheetIfNeeded(_inputPackage);
        }

        #endregion

        #region シート同期

        private void SyncPackageSheets(
            ExcelPackage package, string fileName,
            List<string> sheetsToDelete,
            Dictionary<string, string> renameMap,
            List<(string newName, string templateName)> sheetsToAdd,
            bool isInputFile)
        {
            Logger.Info($"{fileName} のシート同期処理を開始します。");

            // 削除
            foreach (var sheetName in sheetsToDelete)
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
                if (ws != null) { package.Workbook.Worksheets.Delete(ws); Logger.Info($"{fileName}: シート削除 -> {sheetName}"); }
            }

            // リネーム
            foreach (var kvp in renameMap)
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == kvp.Key);
                if (ws != null)
                {
                    ws.Name = kvp.Value;
                    if (isInputFile) UpdateSheetCells(ws);
                    Logger.Info($"{fileName}: シート名変更 -> {kvp.Key} から {kvp.Value}");
                }
            }

            // 追加
            foreach (var (newName, templateName) in sheetsToAdd)
            {
                var templateWs = ResolveTemplateWorksheet(package, newName, templateName, isInputFile);
                if (templateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateName}' が Template.xlsx に見つかりません。");

                int insertIndex = GetInsertIndex(package, newName);
                var newWs = package.Workbook.Worksheets.Add(newName, templateWs);

                if (package.Workbook.Worksheets.Count > 1)
                {
                    try { package.Workbook.Worksheets.MoveAfter(newWs.Index, insertIndex); }
                    catch { Logger.Warn($"{fileName}: シート移動に失敗しました -> {newName}"); }
                }

                UpdateFormulas(newWs, templateName, newName);

                if (isInputFile)
                    SetupNewSheetCells(newWs, newName, templateName);

                Logger.Info($"{fileName}: シート追加 -> {newWs.Name} (テンプレート: {templateName})");
            }

            try { ReorderVehicleSheets(package); }
            catch (Exception ex) { Logger.Warn(ex, "追加後の自動並べ替えに失敗しました。"); }
        }

        private ExcelWorksheet ResolveTemplateWorksheet(ExcelPackage package, string newName, string templateName, bool isInputFile)
        {
            if (!isInputFile)
                return _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);

            try
            {
                bool isEastCeremony   = newName.Contains("東日本セレモニー") || templateName.Contains("東日本セレモニー");
                bool isChFujiYoshida  = newName.Contains("CH富士吉田") || templateName.Contains("CH富士吉田");
                bool isChOotsuki      = newName.Contains("CH大月") || templateName.Contains("CH大月");
                bool isChHigashiFuji  = newName.Contains("CH東富士") || templateName.Contains("CH東富士");

                string preferredTemplate;
                if (isEastCeremony)
                    preferredTemplate = "Template2";
                else if (isChFujiYoshida || isChOotsuki || isChHigashiFuji)
                    preferredTemplate = "Template1";
                else
                {
                    var (branch, _) = ParseSheetNameToBranchAndNumber(newName);
                    bool isEast = (branch?.Contains("東日本") ?? false) || newName.Contains("東日本") || templateName.Contains("東日本");
                    preferredTemplate = isEast ? "Template2" : "Template1";
                }

                return _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == preferredTemplate)
                    ?? _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, $"テンプレート選択で例外が発生しました。既定のテンプレート '{templateName}' を使用します。");
                return _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
            }
        }

        private void SetupNewSheetCells(ExcelWorksheet newWs, string newName, string templateName)
        {
            var (branch, number) = ParseSheetNameToBranchAndNumber(newName);
            if (string.IsNullOrWhiteSpace(branch))
                branch = ParseSheetNameToBranchAndNumber(templateName).Branch;

            var categoryKey         = GetCategoryKey(newName);
            if (categoryKey == "その他") categoryKey = GetCategoryKey(templateName);

            var templateCategory    = GetCategoryKey(templateName);
            string inferredType     = (templateCategory == "霊柩車" || templateCategory == "寝台車") ? templateCategory : null;
            string typeText         = inferredType ?? (categoryKey == "霊柩車" ? "霊柩車" : "寝台車");
            bool branchContainsType = !string.IsNullOrWhiteSpace(branch) && (branch.Contains("霊柩車") || branch.Contains("寝台車"));
            var branchClean         = (branch ?? "").Replace("霊柩車", "").Replace("寝台車", "").Trim();
            string typeTextToUse    = branchContainsType ? (string.IsNullOrWhiteSpace(branchClean) ? typeText : "") : typeText;

            bool isEast = (categoryKey?.Contains("東日本") ?? false)
                || (!string.IsNullOrWhiteSpace(branch) && branch.Contains("東日本"))
                || newName.Contains("東日本");

            if (isEast)
            {
                if (!string.IsNullOrWhiteSpace(branch)) newWs.Cells["B4"].Value = branch;
                if (!string.IsNullOrWhiteSpace(number))
                {
                    if (int.TryParse(number, out int n)) newWs.Cells["C4"].Value = n;
                    else newWs.Cells["C4"].Value = number;
                }
            }
            else
            {
                var d1 = string.IsNullOrWhiteSpace(typeTextToUse)
                    ? branchClean
                    : (string.IsNullOrWhiteSpace(branchClean) ? typeTextToUse : $"{branchClean} {typeTextToUse}").Trim();

                if (!string.IsNullOrWhiteSpace(d1)) newWs.Cells["D1"].Value = d1;

                if (!string.IsNullOrWhiteSpace(number))
                {
                    if (int.TryParse(number, out int n)) newWs.Cells["H1"].Value = n;
                    else newWs.Cells["H1"].Value = number;
                }
            }

            // シート名を「支社名 種類 車両番号」に変更
            var nameParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(branchClean)) nameParts.Add(branchClean);
            if (!string.IsNullOrWhiteSpace(typeTextToUse)) nameParts.Add(typeTextToUse);
            if (!string.IsNullOrWhiteSpace(number)) nameParts.Add(number);
            var finalName = string.Join(" ", nameParts).Trim();
            if (!string.IsNullOrWhiteSpace(finalName))
            {
                try { newWs.Name = finalName; }
                catch (Exception ex) { Logger.Warn(ex, $"シート名変更に失敗しました: {finalName}"); }
            }

            UpdateSheetCells(newWs);
        }

        #endregion

        #region 並べ替え

        private void ReorderVehicleSheets(ExcelPackage package)
        {
            try
            {
                var monthly = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "月間集計");

                var orderedNames = package.Workbook.Worksheets
                    .Where(ws => GetCategoryKey(ws.Name) != "その他" && !ws.Name.Contains("登録")
                              && ws.Name != "月間集計" && !IsTemplateSheet(ws.Name))
                    .Select(ws =>
                    {
                        var (branch, number) = ParseSheetNameToBranchAndNumber(ws.Name);
                        int num = int.TryParse(number, out var n) ? n : int.MaxValue;
                        return new { ws.Name, Branch = branch ?? "", NumberInt = num, CategoryOrder = GetCategoryOrder(ws.Name) };
                    })
                    .OrderBy(v => v.CategoryOrder)
                    .ThenBy(v => v.NumberInt)
                    .ThenBy(v => v.Name, StringComparer.Ordinal)
                    .Select(v => v.Name)
                    .ToList();

                if (monthly != null)
                    package.Workbook.Worksheets.MoveBefore(monthly.Index, 1);

                for (int i = orderedNames.Count - 1; i >= 0; i--)
                {
                    var ws = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == orderedNames[i]);
                    if (ws == null) continue;
                    if (monthly != null) package.Workbook.Worksheets.MoveAfter(ws.Index, 1);
                    else                 package.Workbook.Worksheets.MoveBefore(ws.Index, 1);
                }

                Logger.Info("パッケージのシート順を支社名毎に並べ替えました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "シート並び替え中にエラーが発生しました。");
            }
        }

        #endregion

        #region セル・数式更新

        private void UpdateFormulas(ExcelWorksheet ws, string oldSheetRef, string newSheetRef)
        {
            if (ws.Dimension == null) return;
            // [No.12修正] シート参照の形式は「'シート名'!」と「シート名!」の2パターンがある。
            // 旧実装はシングルクォートありの形式しか置換しておらず、
            // スペースや特殊文字を含まないシート名（引用符なし参照）が更新されないケースがあった。
            // 両パターンを置換することで確実にシート名変更を数式に反映する。
            string newQuoted = NeedQuotes(newSheetRef) ? $"'{newSheetRef}'!" : $"{newSheetRef}!";
            foreach (var cell in ws.Cells)
            {
                if (string.IsNullOrEmpty(cell.Formula)) continue;
                var formula = cell.Formula;
                // 引用符ありパターン（スペース等を含む旧シート名）
                formula = formula.Replace($"'{oldSheetRef}'!", newQuoted);
                // 引用符なしパターン（スペース等を含まない旧シート名）
                if (!NeedQuotes(oldSheetRef))
                    formula = formula.Replace($"{oldSheetRef}!", newQuoted);
                cell.Formula = formula;
            }
            Logger.Info($"シート '{ws.Name}' の数式を更新しました。");
        }

        private void UpdateSheetCells(ExcelWorksheet ws)
        {
            string sheetName = ws.Name;

            if (sheetName.Contains("東日本セレモニー"))
            {
                var m = Regex.Match(sheetName, @"\d+$");
                if (m.Success && int.TryParse(m.Value, out int n))
                    ws.Cells["C4"].Value = n;
                return;
            }

            var numberMatch = Regex.Match(sheetName, @"\d+");
            int? vehicleNumber = null;
            if (numberMatch.Success && int.TryParse(numberMatch.Value, out int num))
                vehicleNumber = num;

            string d1Value = vehicleNumber.HasValue
                ? Regex.Replace(sheetName, @"\s*\d+$", "").Trim()
                : sheetName;

            ws.Cells["D1"].Value = d1Value;
            ws.Cells["H1"].Value = vehicleNumber;
        }

        #endregion

        #region シート分類・命名ヘルパー

        internal (string Branch, string Number) ParseSheetNameToBranchAndNumber(string sheetName)
        {
            var knownPrefixes = new[] { "東日本セレモニー", "CH富士吉田", "CH大月", "CH東富士" };
            foreach (var prefix in knownPrefixes)
            {
                if (sheetName.Contains(prefix))
                {
                    var m = Regex.Match(sheetName, @"\d+$");
                    return (prefix, m.Success ? m.Value : "");
                }
            }

            if (sheetName.StartsWith("霊柩車") || sheetName.StartsWith("寝台車"))
            {
                var parts = sheetName.Split(' ');
                return parts.Length > 1 && int.TryParse(parts.Last(), out _)
                    ? (parts[0], parts.Last())
                    : (parts[0], "");
            }

            return (sheetName, "");
        }

        private int GetCategoryOrder(string sheetName)
        {
            if (sheetName.Contains("CH富士吉田")) return 2;
            if (sheetName.Contains("CH大月") || (sheetName.Contains("大月") && !sheetName.Contains("CH富士吉田") && !sheetName.Contains("CH東富士") && !sheetName.Contains("東日本"))) return 3;
            if (sheetName.Contains("CH東富士")) return 4;
            if (sheetName.Contains("東日本")) return 5;
            if (sheetName.StartsWith("霊柩車") || sheetName.StartsWith("寝台車")) return 1;
            return 99;
        }

        private string GetCategoryKey(string sheetName)
        {
            if (sheetName.Contains("CH富士吉田")) return "CH富士吉田";
            if (sheetName.Contains("CH大月")) return "CH大月";
            if (sheetName.Contains("大月") && !sheetName.Contains("CH富士吉田") && !sheetName.Contains("CH東富士") && !sheetName.Contains("東日本")) return "CH大月";
            if (sheetName.Contains("CH東富士")) return "CH東富士";
            if (sheetName.Contains("東日本セレモニー")) return "東日本セレモニー";
            if (sheetName.StartsWith("霊柩車")) return "通常-霊柩車";
            if (sheetName.StartsWith("寝台車")) return "通常-寝台車";
            return "その他";
        }

        private int GetInsertIndex(ExcelPackage package, string newSheetName)
        {
            int newOrder  = GetCategoryOrder(newSheetName);
            int lastIndex = 0;
            foreach (var ws in package.Workbook.Worksheets.OrderBy(s => s.Index))
                if (GetCategoryOrder(ws.Name) <= newOrder) lastIndex = ws.Index;
            return lastIndex > 0 ? lastIndex : package.Workbook.Worksheets.Count;
        }

        #endregion
    }
}
