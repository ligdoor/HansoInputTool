using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace HansoInputTool.Services
{
    public class ExcelHandler
    {
        static ExcelHandler()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly string _inputFilePath;
        private readonly string _templateFilePath;
        private readonly ColumnMapping _columnMap;
        private ExcelPackage _inputPackage;
        private ExcelPackage _templatePackage;
        private readonly Dictionary<string, List<RowData>> _dataCache = new();

        public ExcelHandler(string inputFilePath, string templateFilePath, ColumnMapping columnMap)
        {
            _inputFilePath = inputFilePath;
            _templateFilePath = templateFilePath;
            _columnMap = columnMap;
            Load();
        }
        private bool NeedQuotes(string sheetName)
        {
            return sheetName.Contains(" ") ||
                   sheetName.Contains("-") ||
                   sheetName.Contains("(") ||
                   sheetName.Contains(")") ||
                   sheetName.Contains("'") ||
                   sheetName.Contains("!") ||
                   sheetName.Contains("#");
        }

        public void Load()
        {
            _inputPackage?.Dispose();
            _templatePackage?.Dispose();
            _inputPackage = new ExcelPackage(new FileInfo(_inputFilePath));
            _templatePackage = new ExcelPackage(new FileInfo(_templateFilePath));
            _dataCache.Clear();
        }

        // Save の誤字修正
        public void Save()
        {
            _inputPackage?.Save();
            _templatePackage?.Save();
        }

        public bool TemplateSheetExists(string sheetName)
        {
            return _templatePackage.Workbook.Worksheets.Any(ws => ws.Name == sheetName);
        }
        public List<string> SheetNames => _inputPackage?.Workbook.Worksheets
            .Where(ws => !ws.Name.Contains("登録") && !IsTemplateSheet(ws.Name))
            .Select(ws => ws.Name)
            .ToList() ?? new List<string>();

        public void SyncAllVehicleSheets(List<string> sheetsToDelete, Dictionary<string, string> renameMap, List<(string newName, string templateName)> sheetsToAdd)
        {
            SyncPackageSheets(_inputPackage, "Input.xlsx", sheetsToDelete, renameMap, sheetsToAdd, true);
            SyncPackageSheets(_templatePackage, "Template.xlsx", sheetsToDelete, renameMap, sheetsToAdd, false);

            // 両ファイルともシートを支社名毎に並べ替え
            ReorderVehicleSheets(_inputPackage);
            ReorderVehicleSheets(_templatePackage);

            // 月間集計は Input 側を更新する（Template を破壊しない）
            UpdateMonthlySummarySheetIfNeeded(_inputPackage);
        }

        private void SyncPackageSheets(ExcelPackage package, string fileName, List<string> sheetsToDelete, Dictionary<string, string> renameMap, List<(string newName, string templateName)> sheetsToAdd, bool isInputFile)
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
                ExcelWorksheet templateWs;
                if (isInputFile)
                {
                    // 明示ルール: CH系は Template1、東日本セレモニーは Template2
                    try
                    {
                        var resolved = ParseSheetNameToBranchAndNumber(newName);
                        var branch = resolved.Branch ?? "";

                        bool isChFujiYoshida = newName.Contains("CH富士吉田") || templateName.Contains("CH富士吉田");
                        bool isChOotsuki = newName.Contains("CH大月") || templateName.Contains("CH大月");
                        bool isChHigashiFuji = newName.Contains("CH東富士") || templateName.Contains("CH東富士");
                        bool isEastCeremony = newName.Contains("東日本セレモニー") || templateName.Contains("東日本セレモニー") || branch.Contains("東日本");

                        string preferredTemplate;
                        if (isEastCeremony)
                        {
                            preferredTemplate = "Template2";
                        }
                        else if (isChFujiYoshida || isChOotsuki || isChHigashiFuji)
                        {
                            preferredTemplate = "Template1";
                        }
                        else
                        {
                            bool isEast = branch.Contains("東日本") || newName.Contains("東日本") || templateName.Contains("東日本");
                            preferredTemplate = isEast ? "Template2" : "Template1";
                        }

                        templateWs = _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == preferredTemplate)
                                     ?? _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(ex, $"テンプレート選択で例外が発生しました。既定のテンプレート '{templateName}' を使用します。");
                        templateWs = _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
                    }
                }
                else
                {
                    templateWs = _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
                }

                if (templateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateName}' が Template.xlsx に見つかりません。");

                // 挿入位置は追加されるシート名（newName）に基づいて決定する
                int insertIndex = GetInsertIndex(package, newName);
                var newWs = package.Workbook.Worksheets.Add(newName, templateWs);

                if (package.Workbook.Worksheets.Count > 1)
                {
                    try
                    {
                        package.Workbook.Worksheets.MoveAfter(newWs.Index, insertIndex);
                    }
                    catch
                    {
                        Logger.Warn($"{fileName}: シート移動に失敗しました -> {newName}");
                    }
                }

                UpdateFormulas(newWs, templateName, newName);

                if (isInputFile)
                {
                    // 支社名・番号・種類を解決
                    var resolved = ParseSheetNameToBranchAndNumber(newName);
                    var branch = resolved.Branch;
                    var number = resolved.Number;

                    if (string.IsNullOrWhiteSpace(branch))
                    {
                        branch = ParseSheetNameToBranchAndNumber(templateName).Branch;
                    }

                    // 種類はカテゴリキーで判定（寝台車 or 霊柩車）
                    var categoryKey = GetCategoryKey(newName);
                    if (categoryKey == "その他") categoryKey = GetCategoryKey(templateName);

                    // テンプレート側に明示的な種類情報があれば優先して利用する
                    var templateCategory = GetCategoryKey(templateName);
                    string inferredTypeFromTemplate = templateCategory == "霊柩車" || templateCategory == "寝台車" ? templateCategory : null;

                    // 優先ルール: テンプレート側の種類 -> シート名のカテゴリ -> デフォルト寝台車
                    string typeText = inferredTypeFromTemplate ?? (categoryKey == "霊柩車" ? "霊柩車" : "寝台車");

                    // branch に既にタイプが含まれているかチェック
                    bool branchContainsType = !string.IsNullOrWhiteSpace(branch) && (branch.Contains("霊柩車") || branch.Contains("寝台車"));
                    // branch から種類語を取り除いた支社名（空なら ""）
                    var branchClean = (branch ?? "").Replace("霊柩車", "").Replace("寝台車", "").Trim();

                    // typeText を最終的に使うか決める
                    string typeTextToUse;
                    if (!branchContainsType) typeTextToUse = typeText;
                    else typeTextToUse = string.IsNullOrWhiteSpace(branchClean) ? typeText : "";

                    // 東日本系は B4/C4 に設定（Template2 相当）
                    bool isEast = (categoryKey != null && categoryKey.Contains("東日本")) || (!string.IsNullOrWhiteSpace(branch) && branch.Contains("東日本")) || newName.Contains("東日本");
                    if (isEast)
                    {
                        if (!string.IsNullOrWhiteSpace(branch)) newWs.Cells["B4"].Value = branch;
                        if (!string.IsNullOrWhiteSpace(number) && int.TryParse(number, out int n1)) newWs.Cells["C4"].Value = n1;
                        else if (!string.IsNullOrWhiteSpace(number)) newWs.Cells["C4"].Value = number;
                    }
                    else
                    {
                        // D1 に表示する支社名（種類語を除いた branchClean を優先）
                        var d1 = string.IsNullOrWhiteSpace(typeTextToUse)
                            ? branchClean
                            : (string.IsNullOrWhiteSpace(branchClean) ? typeTextToUse : $"{branchClean} {typeTextToUse}").Trim();

                        if (!string.IsNullOrWhiteSpace(d1)) newWs.Cells["D1"].Value = d1;

                        if (!string.IsNullOrWhiteSpace(number) && int.TryParse(number, out int n2)) newWs.Cells["H1"].Value = n2;
                        else if (!string.IsNullOrWhiteSpace(number)) newWs.Cells["H1"].Value = number;
                    }

                    // シート名を「支社名 種類 車両番号」に変更（重複を避け、branch が種類語のみなら種類語は付与する）
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

                    // B4/C4 の上書きに備え補助処理を呼ぶ
                    UpdateSheetCells(newWs);
                }

                Logger.Info($"{fileName}: シート追加 -> {newWs.Name} (テンプレート: {templateName})");
            }

            // 追加後は必ず並べ替え（追加したシートも含める）
            try
            {
                ReorderVehicleSheets(package);
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "追加後の自動並べ替えに失敗しました。");
            }
        }

        private void UpdateMonthlySummarySheetIfNeeded(ExcelPackage package)
        {
            var summarySheet = package.Workbook.Worksheets["月間集計"];
            if (summarySheet == null)
            {
                Logger.Warn("月間集計シートが見つかりません。");
                return;
            }

            Logger.Info("=== 月間集計シートの更新を開始 ===");

            // 対象シート一覧を取得（月間集計シート自体を除外）
            var allVehicleSheets = package.Workbook.Worksheets
                .Where(ws => ws.Name != "月間集計")
                .Select(ws => ws.Name)
                .OrderBy(s => GetCategoryOrder(s))
                .ThenBy(s => s)
                .ToList();

            Logger.Info($"対象車両シート数: {allVehicleSheets.Count}");

            // 固定値（月間集計シートの構造に基づく）
            int dataStartRow = 6;
            int startCol = 1;  // A列
            int endCol = 11;   // K列
            int maxDataRows = 69; // テーブルの最大データ行数

            // 既存データをすべてクリア
            for (int row = dataStartRow; row < dataStartRow + maxDataRows; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = summarySheet.Cells[row, col];
                    cell.Value = null;
                    cell.Formula = null;
                }
            }
            Logger.Info($"{maxDataRows}行分のデータをクリアしました");

            // 新しいデータを書き込み
            for (int i = 0; i < allVehicleSheets.Count; i++)
            {
                string sheetName = allVehicleSheets[i];
                var (branch, number) = ParseSheetNameToBranchAndNumber(sheetName);
                int currentRow = dataStartRow + i;

                try
                {
                    // A列: No.
                    summarySheet.Cells[currentRow, 1].Value = $"No.{i + 1}";

                    // B列: 営業所
                    summarySheet.Cells[currentRow, 2].Value = branch;

                    // C列: 番号
                    summarySheet.Cells[currentRow, 3].Value = int.TryParse(number, out int num) ? num : (object)number;

                    // シート名に特殊文字が含まれる場合はシングルクォートで囲む
                    string safeSheetName = NeedQuotes(sheetName) ? $"'{sheetName}'" : sheetName;

                    // D列: 稼働日数（参照: 各シートのE4）
                    summarySheet.Cells[currentRow, 4].Formula = $"{safeSheetName}!E4";

                    // E列: 搬送回数（参照: 各シートのG4）
                    summarySheet.Cells[currentRow, 5].Formula = $"{safeSheetName}!G4";

                    // F列: 平均km（計算: D列/E列）
                    summarySheet.Cells[currentRow, 6].Formula = $"IF(E{currentRow}>0,D{currentRow}/E{currentRow},0)";

                    // G列: 搬送回数（参照: 各シートのG4）※E列と同じ
                    summarySheet.Cells[currentRow, 7].Formula = $"{safeSheetName}!G4";

                    // H列: 有料km（参照: 各シートのH4）
                    summarySheet.Cells[currentRow, 8].Formula = $"{safeSheetName}!H4";

                    // I列: 無料km（参照: 各シートのI4）
                    summarySheet.Cells[currentRow, 9].Formula = $"{safeSheetName}!I4";

                    // J列: 合計km（計算: H列+I列）
                    summarySheet.Cells[currentRow, 10].Formula = $"H{currentRow}+I{currentRow}";

                    // K列: 金額合計（参照: 各シートのK4）
                    summarySheet.Cells[currentRow, 11].Formula = $"{safeSheetName}!K4";

                    Logger.Info($"Row {currentRow}: {sheetName} のデータを設定しました");
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"Row {currentRow} ({sheetName}) のデータ設定中にエラーが発生しました");
                    // エラーが発生しても処理を継続
                }
            }

            // データがない場合
            if (!allVehicleSheets.Any())
            {
                summarySheet.Cells[dataStartRow, 1].Value = "（車両データなし）";
                Logger.Info("対象車両シートが0件のため、メッセージを表示しました");
            }

            // 計算モードを自動に設定
            package.Workbook.CalcMode = ExcelCalcMode.Automatic;

            Logger.Info("=== 月間集計シートの更新が完了しました ===");
        }
        // Template 判定
        private static bool IsTemplateSheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            var normalized = sheetName.Replace(" ", "").ToLowerInvariant();
            if (normalized.StartsWith("template", StringComparison.OrdinalIgnoreCase)) return true;
            if (sheetName.IndexOf("テンプレート", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        // ================================================================
        // 以降：元の機能群（復元済み）
        // ================================================================
        public List<string> GetVehicleSheetNames()
        {
            return _inputPackage.Workbook.Worksheets
                .Where(s => !s.Name.Contains("登録") && !IsTemplateSheet(s.Name))
                .Select(s => s.Name)
                .ToList();
        }

        // 安定して並べ替える実装（逆順で先頭へ移動）
        private void ReorderVehicleSheets(ExcelPackage package)
        {
            try
            {
                var monthly = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "月間集計");

                var vehicleInfos = package.Workbook.Worksheets
                    .Where(ws => GetCategoryKey(ws.Name) != "その他" && !ws.Name.Contains("登録") && ws.Name != "月間集計" && !IsTemplateSheet(ws.Name))
                    .Select(ws =>
                    {
                        var (branch, number) = ParseSheetNameToBranchAndNumber(ws.Name);
                        int num = int.TryParse(number, out var n) ? n : int.MaxValue;
                        return new { Name = ws.Name, Branch = branch ?? "", NumberInt = num, CategoryOrder = GetCategoryOrder(ws.Name) };
                    })
                    .ToList();

                var orderedNames = vehicleInfos
                    .OrderBy(v => v.Branch, StringComparer.Ordinal)
                    .ThenBy(v => v.CategoryOrder)
                    .ThenBy(v => v.NumberInt)
                    .ThenBy(v => v.Name, StringComparer.Ordinal)
                    .Select(v => v.Name)
                    .ToList();

                // 逆順で先頭に移動
                for (int i = orderedNames.Count - 1; i >= 0; i--)
                {
                    var name = orderedNames[i];
                    var ws = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == name);
                    if (ws == null) continue;
                    package.Workbook.Worksheets.MoveBefore(ws.Index, 1);
                }

                if (monthly != null)
                {
                    package.Workbook.Worksheets.MoveAfter(monthly.Index, package.Workbook.Worksheets.Count);
                }

                Logger.Info($"パッケージのシート順を支社名毎に並べ替えました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "シート並び替え中にエラーが発生しました。");
            }
        }

        private void UpdateFormulas(ExcelWorksheet ws, string oldSheetRef, string newSheetRef)
        {
            if (ws.Dimension == null) return;
            foreach (var cell in ws.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    cell.Formula = cell.Formula.Replace($"'{oldSheetRef}'!", $"'{newSheetRef}'!");
                }
            }
            Logger.Info($"シート '{ws.Name}' の数式を更新しました。");
        }

        private (string Branch, string Number) ParseSheetNameToBranchAndNumber(string sheetName)
        {
            // 東日本セレモニーの場合
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("東日本セレモニー", numberMatch.Success ? numberMatch.Value : "");
            }

            // CH富士吉田の場合
            if (sheetName.Contains("CH富士吉田"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("CH富士吉田", numberMatch.Success ? numberMatch.Value : "");
            }

            // CH大月の場合
            if (sheetName.Contains("CH大月"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("CH大月", numberMatch.Success ? numberMatch.Value : "");
            }

            // CH東富士の場合
            if (sheetName.Contains("CH東富士"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("CH東富士", numberMatch.Success ? numberMatch.Value : "");
            }

            // 通常の車両（営業所名なし） - 霊柩車または寝台車
            if (sheetName.StartsWith("霊柩車") || sheetName.StartsWith("寝台車"))
            {
                var parts = sheetName.Split(' ');
                if (parts.Length > 1 && int.TryParse(parts.Last(), out _))
                {
                    return (parts[0], parts.Last());
                }
                // 番号がない場合
                return (parts[0], "");
            }

            return (sheetName, "");
        }
        private (string Branch, string Number) ParseSheetNameToBranchAndNumberForNormalSheet(string sheetName)
        {
            var parts = sheetName.Split(' ');
            if (parts.Length > 1 && int.TryParse(parts.Last(), out _)) { return (string.Join(" ", parts.Take(parts.Length - 1)), parts.Last()); }
            return (sheetName, "");
        }

        private int GetCategoryOrder(string sheetName)
        {
            // 通常の車両（営業所名なし） - 最優先
            if (!sheetName.Contains("CH富士吉田") &&
                !sheetName.Contains("CH大月") &&
                !sheetName.Contains("CH東富士") &&
                !sheetName.Contains("東日本") &&
                (sheetName.StartsWith("霊柩車") || sheetName.StartsWith("寝台車")))
            {
                return 1;
            }

            // CH富士吉田
            if (sheetName.Contains("CH富士吉田"))
                return 2;

            // CH大月
            if (sheetName.Contains("CH大月"))
                return 3;

            // CH東富士
            if (sheetName.Contains("CH東富士"))
                return 4;

            // 東日本セレモニー
            if (sheetName.Contains("東日本"))
                return 5;

            return 99;
        }
        private string GetCategoryKey(string sheetName)
        {
            // 営業所名が明示されている場合
            if (sheetName.Contains("CH富士吉田")) return "CH富士吉田";
            if (sheetName.Contains("CH大月")) return "CH大月";
            if (sheetName.Contains("CH東富士")) return "CH東富士";
            if (sheetName.Contains("東日本セレモニー")) return "東日本セレモニー";

            // 営業所の指定がない通常の車両
            if (sheetName.StartsWith("霊柩車")) return "通常-霊柩車";
            if (sheetName.StartsWith("寝台車")) return "通常-寝台車";

            return "その他";
        }
        private int GetInsertIndex(ExcelPackage package, string newSheetName)
        {
            var newCategoryOrder = GetCategoryOrder(newSheetName);
            int lastIndex = 0;
            foreach (var ws in package.Workbook.Worksheets.OrderBy(s => s.Index))
            {
                int order = GetCategoryOrder(ws.Name);
                if (order <= newCategoryOrder) lastIndex = ws.Index;
            }
            return lastIndex > 0 ? lastIndex : package.Workbook.Worksheets.Count;
        }

        private static int FindTotalRow(ExcelWorksheet ws)
        {
            if (ws?.Dimension == null) return -1;
            for (int row = ws.Dimension.End.Row; row >= 3; row--)
            {
                if (ws.Cells[row, 1].Value?.ToString()?.Contains("合計") == true) return row;
            }
            return -1;
        }

        private static int? GetNullableInt(object val)
        {
            if (val == null) return null;
            if (val is int i) return i;
            if (val is long l) return (int)l;
            if (val is double d) return (int)d;
            if (val is decimal m) return (int)m;

            var s = val.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            s = s.Replace(",", "").Replace("，", "");

            if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double parsed) ||
                double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed))
            {
                return (int)parsed;
            }

            Logger.Warn($"GetNullableInt: 非数値フィールドをパースできませんでした: '{s}'");
            return null;
        }

        private static double? GetNullableDouble(object val) => val == null ? null : Convert.ToDouble(val);

        public List<RowData> GetSheetDataForPreview(string sheetName)
        {
            if (sheetName == null || !_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) return new List<RowData>();
            if (_dataCache.ContainsKey(sheetName)) return _dataCache[sheetName];

            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return new List<RowData>();

            var data = new List<RowData>();
            var map = _columnMap.NormalSheet;
            bool isOotsuki = sheetName.Contains("大月");

            for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
            {
                if (ws.Cells[rowIndex, map.Day].Value == null && ws.Cells[rowIndex, map.YuryoKm].Value == null) continue;

                var rowData = new RowData
                {
                    RowIndex = rowIndex,
                    B_Day = GetNullableInt(ws.Cells[rowIndex, map.Day].Value),
                    C_Hanso = GetNullableInt(ws.Cells[rowIndex, map.HansoCount].Value),
                    D_YuryoKm = GetNullableInt(ws.Cells[rowIndex, map.YuryoKm].Value),
                    E_MuryoKm = GetNullableInt(ws.Cells[rowIndex, map.MuryoKm].Value),
                    H_LateFeeOotsuki = GetNullableInt(ws.Cells[rowIndex, map.ShinyaFee].Value),
                    K_LateMinutes = GetNullableInt(ws.Cells[rowIndex, map.ShinyaMinutes].Value),
                    L_IsKoryo = GetNullableInt(ws.Cells[rowIndex, map.IsKoryo].Value)
                };

                rowData.LateValueText = isOotsuki ? rowData.H_LateFeeOotsuki?.ToString() : rowData.K_LateMinutes?.ToString();
                data.Add(rowData);
            }

            _dataCache[sheetName] = data;
            return data;
        }

        private void UpdateSheetCells(ExcelWorksheet ws)
        {
            string sheetName = ws.Name;

            // 東日本セレモニーの場合のみ特殊処理（C4に番号）
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                if (numberMatch.Success && int.TryParse(numberMatch.Value, out int number))
                {
                    ws.Cells["C4"].Value = number;
                }
                return;
            }

            // その他の車両はすべてD1とH1に設定

            // 番号を抽出
            var numberMatch2 = Regex.Match(sheetName, @"\d+");
            int? vehicleNumber = null;
            if (numberMatch2.Success && int.TryParse(numberMatch2.Value, out int num))
            {
                vehicleNumber = num;
            }

            // D1に営業所名+車種名を設定（番号以外の部分）
            string d1Value = sheetName;
            if (vehicleNumber.HasValue)
            {
                // 最後の数字を削除
                d1Value = Regex.Replace(sheetName, @"\s*\d+$", "").Trim();
            }

            ws.Cells["D1"].Value = d1Value;

            // H1に番号を設定
            ws.Cells["H1"].Value = vehicleNumber;
        }

        // 通常シート登録（行挿入含む）: (targetRow, insertInfo) を返す
        public (int targetRow, string insertInfo) RegisterNormalData(string sheetName, Dictionary<string, double?> values, bool isKoryo)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) throw new Exception($"シート '{sheetName}' に '合計' 行が見つかりません。");

            var map = _columnMap.NormalSheet;
            int targetRow = -1;
            for (int r = 3; r < totalRowIndex; r++)
            {
                if (ws.Cells[r, map.Day].Value == null) { targetRow = r; break; }
            }
            string insertInfo = "";
            if (targetRow == -1)
            {
                ws.InsertRow(totalRowIndex, 1);
                targetRow = totalRowIndex;
                insertInfo = "空き行がないため、合計行の上に新しい行を挿入します。";
            }

            // 値設定
            double? yuryoVal = values.GetValueOrDefault("有料キロ(D)");
            int hansoVal = (yuryoVal.HasValue && yuryoVal > 0) ? 1 : 0;
            ws.Cells[targetRow, map.Day].Value = values.GetValueOrDefault("日(B)");
            ws.Cells[targetRow, map.HansoCount].Value = hansoVal;
            ws.Cells[targetRow, map.YuryoKm].Value = yuryoVal;
            ws.Cells[targetRow, map.MuryoKm].Value = values.GetValueOrDefault("無料キロ(E)");
            ws.Cells[targetRow, map.IsKoryo].Value = isKoryo ? 1 : (object)null;

            bool isOotsuki = sheetName.Contains("大月");
            if (isOotsuki)
            {
                ws.Cells[targetRow, map.ShinyaFee].Value = values.GetValueOrDefault("深夜料金(H)");
                ws.Cells[targetRow, map.ShinyaMinutes].Value = null;
            }
            else
            {
                ws.Cells[targetRow, map.ShinyaFee].Value = null;
                ws.Cells[targetRow, map.ShinyaMinutes].Value = values.GetValueOrDefault("深夜時間(K)");
            }

            // ★修正点：登録後にキャッシュを削除してプレビューが正しく更新されるようにする
            if (_dataCache.ContainsKey(sheetName)) _dataCache.Remove(sheetName);

            return (targetRow, insertInfo);
        }

        // 通常シート行更新
        public void UpdateNormalData(string sheetName, int rowIndex, Dictionary<string, double?> values, bool isKoryo)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.NormalSheet;

            double? yuryoVal = values.GetValueOrDefault("有料キロ(D)");
            int hansoVal = (yuryoVal.HasValue && yuryoVal > 0) ? 1 : 0;
            ws.Cells[rowIndex, map.Day].Value = values.GetValueOrDefault("日(B)");
            ws.Cells[rowIndex, map.HansoCount].Value = hansoVal;
            ws.Cells[rowIndex, map.YuryoKm].Value = yuryoVal;
            ws.Cells[rowIndex, map.MuryoKm].Value = values.GetValueOrDefault("無料キロ(E)");
            ws.Cells[rowIndex, map.IsKoryo].Value = isKoryo ? 1 : (object)null;

            bool isOotsuki = sheetName.Contains("大月");
            if (isOotsuki)
            {
                ws.Cells[rowIndex, map.ShinyaFee].Value = values.GetValueOrDefault("深夜料金(H)");
                ws.Cells[rowIndex, map.ShinyaMinutes].Value = null;
            }
            else
            {
                ws.Cells[rowIndex, map.ShinyaFee].Value = null;
                ws.Cells[rowIndex, map.ShinyaMinutes].Value = values.GetValueOrDefault("深夜時間(K)");
            }

            // ★修正点：編集後にキャッシュを削除してプレビューが正しく更新されるようにする
            if (_dataCache.ContainsKey(sheetName)) _dataCache.Remove(sheetName);
        }

        // 東日本シートの登録
        public void RegisterEastData(string sheetName, Dictionary<string, double?> values)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.EastSheet;
            ws.Cells[map.Jitsudo].Value = values.GetValueOrDefault("延実働車輌数");
            ws.Cells[map.Hanso].Value = values.GetValueOrDefault("搬送回数");
            ws.Cells[map.YuryoKm].Value = values.GetValueOrDefault("有料キロ数");
            ws.Cells[map.MuryoKm].Value = values.GetValueOrDefault("無料キロ数");
            ws.Cells[map.UnsoJisseki].Value = values.GetValueOrDefault("運輸実績");
        }

        // 行削除
        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r)) { ws.DeleteRow(rowIndex); }
            // ★行削除後もキャッシュをクリア
            if (_dataCache.ContainsKey(sheetName)) _dataCache.Remove(sheetName);
        }

        // 全データクリア（ログメッセージのリストを返す）
        public List<string> ClearData()
        {
            var logMessages = new List<string>();
            var normalMap = _columnMap.NormalSheet;
            var eastMap = _columnMap.EastSheet;
            foreach (var ws in _inputPackage.Workbook.Worksheets)
            {
                if (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                {
                    var totalRowIndex = FindTotalRow(ws);
                    if (totalRowIndex != -1)
                    {
                        for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
                        {
                            foreach (int colIndex in new[] { normalMap.Day, normalMap.HansoCount, normalMap.YuryoKm, normalMap.MuryoKm, normalMap.ShinyaFee, normalMap.ShinyaMinutes, normalMap.IsKoryo })
                            {
                                if (colIndex > 0) ws.Cells[rowIndex, colIndex].Value = null;
                            }
                        }
                        logMessages.Add($"[{ws.Name}] の入力値をクリアしました。");
                    }
                }
                else if (ws.Name.Contains("東日本"))
                {
                    ws.Cells[eastMap.Jitsudo].Value = null;
                    ws.Cells[eastMap.Hanso].Value = null;
                    ws.Cells[eastMap.YuryoKm].Value = null;
                    ws.Cells[eastMap.MuryoKm].Value = null;
                    ws.Cells[eastMap.UnsoJisseki].Value = null;
                    logMessages.Add($"[{ws.Name}] のデータをクリアしました。");
                }
            }
            _dataCache.Clear();
            return logMessages;
        }

        // 残データチェック
        public bool CheckRemainingData()
        {
            var map = _columnMap.NormalSheet;
            foreach (var ws in _inputPackage.Workbook.Worksheets)
            {
                if ((ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH")) && ws.Cells[3, map.Day].Value != null)
                    return true;
            }
            return false;
        }
    }
}