using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HansoInputTool.Services
{
    public class ExcelHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly string _inputFilePath;
        private readonly string _templateFilePath;
        private readonly ColumnMapping _columnMap;
        private ExcelPackage _inputPackage;
        private ExcelPackage _templatePackage;
        private readonly Dictionary<string, List<RowData>> _dataCache = new();

        public List<string> SheetNames => _inputPackage?.Workbook.Worksheets.Select(ws => ws.Name).ToList() ?? new List<string>();

        public ExcelHandler(string inputFilePath, string templateFilePath, ColumnMapping columnMap)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _inputFilePath = inputFilePath;
            _templateFilePath = templateFilePath;
            _columnMap = columnMap;
            Load();
        }

        public void Load()
        {
            _inputPackage?.Dispose();
            _templatePackage?.Dispose();
            _inputPackage = new ExcelPackage(new FileInfo(_inputFilePath));
            _templatePackage = new ExcelPackage(new FileInfo(_templateFilePath));
        }

        public void Save()
        {
            _inputPackage?.Save();
            _templatePackage?.Save();
        }

        public void SyncAllVehicleSheets(List<string> sheetsToDelete, Dictionary<string, string> renameMap, List<(string newName, string templateName)> sheetsToAdd)
        {
            SyncPackageSheets(_inputPackage, "Input.xlsx", sheetsToDelete, renameMap, sheetsToAdd, true);
            SyncPackageSheets(_templatePackage, "Template.xlsx", sheetsToDelete, renameMap, sheetsToAdd, false);
            UpdateMonthlySummarySheetIfNeeded(_templatePackage);
        }

        private void SyncPackageSheets(ExcelPackage package, string fileName, List<string> sheetsToDelete, Dictionary<string, string> renameMap, List<(string newName, string templateName)> sheetsToAdd, bool isInputFile)
        {
            Logger.Info($"{fileName} のシート同期処理を開始します。");

            foreach (var sheetName in sheetsToDelete)
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
                if (ws != null) { package.Workbook.Worksheets.Delete(ws); Logger.Info($"{fileName}: シート削除 -> {sheetName}"); }
            }

            // 名前の変更は、元のシート名を持つシートの Name プロパティを直接変更する
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

            foreach (var (newName, templateName) in sheetsToAdd)
            {
                var templateWs = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == templateName);
                if (templateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateName}' が{fileName}に見つかりません。");

                int insertIndex = GetInsertIndex(package, templateName);
                var newWs = package.Workbook.Worksheets.Copy(templateWs.Name, newName);

                // .Position ではなく .Index を使用する
                if (package.Workbook.Worksheets.Count > 1)
                {
                    package.Workbook.Worksheets.MoveAfter(newWs.Index, insertIndex);
                }

                if (isInputFile) UpdateSheetCells(newWs);
                Logger.Info($"{fileName}: シート追加 -> {newName} (テンプレート: {templateName})");
            }
        }

        private void UpdateMonthlySummarySheetIfNeeded(ExcelPackage package)
        {
            var summarySheet = package.Workbook.Worksheets["月間集計"];
            if (summarySheet == null) return;

            Logger.Info("月間集計シートの同期処理を開始します。");

            var allVehicleSheets = package.Workbook.Worksheets.Where(ws => GetCategoryKey(ws.Name) != "その他" && ws.Name != "月間集計").Select(ws => ws.Name).ToList();

            int totalRow = (summarySheet.Dimension?.End.Row ?? 3);
            while (totalRow >= 4 && (summarySheet.Cells[totalRow, 2].Value == null || string.IsNullOrWhiteSpace(summarySheet.Cells[totalRow, 2].Text)))
            {
                totalRow--;
            }

            if (totalRow >= 4) { summarySheet.DeleteRow(4, totalRow - 3); }

            int currentRow = 4;
            foreach (var sheetName in allVehicleSheets.OrderBy(s => GetCategoryOrder(s)).ThenBy(s => s))
            {
                summarySheet.InsertRow(currentRow, 1, currentRow > 4 ? currentRow - 1 : 3);

                var (branch, number) = ParseSheetNameToBranchAndNumber(sheetName);
                summarySheet.Cells[currentRow, 1].Value = $"No.{currentRow - 3}";
                summarySheet.Cells[currentRow, 2].Value = branch;
                summarySheet.Cells[currentRow, 3].Value = int.TryParse(number, out int num) ? num : (object)number;
                currentRow++;
            }
        }

        public List<string> GetVehicleSheetNames()
        {
            return _inputPackage.Workbook.Worksheets
                .Where(s => !s.Name.Contains("登録"))
                .Select(s => s.Name)
                .ToList();
        }

        // (以下のメソッドは旧バージョンとの互換性のために残すが、基本的には使われない)
        public void AddVehicleSheet(string newSheetName, string templateSheetName)
        {
            Logger.Info($"シート追加処理を開始: {newSheetName} (テンプレート: {templateSheetName})");
            var inputTemplateWs = _inputPackage.Workbook.Worksheets[templateSheetName];
            if (inputTemplateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateSheetName}' がInput.xlsxに見つかりません。");
            int inputIndex = GetInsertIndex(_inputPackage, templateSheetName);
            var newInputWs = _inputPackage.Workbook.Worksheets.Copy(inputTemplateWs.Name, newSheetName);
            if (_inputPackage.Workbook.Worksheets.Count > 1) _inputPackage.Workbook.Worksheets.MoveAfter(newInputWs.Index, inputIndex);
            UpdateSheetCells(newInputWs);
            var templateTemplateWs = _templatePackage.Workbook.Worksheets[templateSheetName];
            if (templateTemplateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateSheetName}' がTemplate.xlsxに見つかりません。");
            int templateIndex = GetInsertIndex(_templatePackage, templateSheetName);
            var newTemplateWs = _templatePackage.Workbook.Worksheets.Copy(templateTemplateWs.Name, newSheetName);
            if (_templatePackage.Workbook.Worksheets.Count > 1) _templatePackage.Workbook.Worksheets.MoveAfter(newTemplateWs.Index, templateIndex);
        }

        public void DeleteVehicleSheet(string sheetName)
        {
            Logger.Info($"シート削除処理を開始: {sheetName}");
            var inputWs = _inputPackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
            if (inputWs != null) _inputPackage.Workbook.Worksheets.Delete(inputWs);
            var templateWs = _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
            if (templateWs != null) _templatePackage.Workbook.Worksheets.Delete(templateWs);
        }

        public void RenameVehicleSheet(string oldSheetName, string newSheetName)
        {
            Logger.Info($"シート名変更処理を開始: {oldSheetName} -> {newSheetName}");
            var inputWs = _inputPackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == oldSheetName);
            if (inputWs != null) { inputWs.Name = newSheetName; UpdateSheetCells(inputWs); }
            var templateWs = _templatePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == oldSheetName);
            if (templateWs != null) { templateWs.Name = newSheetName; }
        }

        private (string Branch, string Number) ParseSheetNameToBranchAndNumber(string sheetName)
        {
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("東日本", numberMatch.Success ? numberMatch.Value : "");
            }
            var parts = sheetName.Split(' ');
            if (parts.Length > 1 && int.TryParse(parts.Last(), out _))
            {
                return (string.Join(" ", parts.Take(parts.Length - 1)), parts.Last());
            }
            return (sheetName, "");
        }

        private int GetCategoryOrder(string sheetName)
        {
            if (sheetName.Contains("CH富士吉田")) return 1;
            if (sheetName.Contains("CH大月")) return 2;
            if (sheetName.Contains("CH東富士")) return 3;
            if (sheetName.Contains("霊柩車")) return 4;
            if (sheetName.Contains("寝台車")) return 5;
            if (sheetName.Contains("東日本")) return 6;
            return 99;
        }

        private string GetCategoryKey(string sheetName)
        {
            if (sheetName.Contains("CH大月")) return "CH大月";
            if (sheetName.Contains("CH東富士")) return "CH東富士";
            if (sheetName.Contains("東日本セレモニー")) return "東日本セレモニー";
            if (sheetName.Contains("霊柩車")) return "霊柩車";
            if (sheetName.Contains("寝台車")) return "寝台車";
            return "その他";
        }

        private int GetInsertIndex(ExcelPackage package, string baseSheetName)
        {
            string categoryKey = GetCategoryKey(baseSheetName);
            var categorySheets = package.Workbook.Worksheets
                .Where(ws => GetCategoryKey(ws.Name) == categoryKey)
                .ToList();
            if (categorySheets.Any()) { return categorySheets.Max(ws => ws.Index); }
            return package.Workbook.Worksheets.Count;
        }

        private void UpdateSheetCells(ExcelWorksheet ws)
        {
            string sheetName = ws.Name;
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                if (numberMatch.Success && int.TryParse(numberMatch.Value, out int number)) { ws.Cells["C4"].Value = number; }
            }
            else
            {
                var lastSpaceIndex = sheetName.LastIndexOf(' ');
                if (lastSpaceIndex > -1 && int.TryParse(sheetName.Substring(lastSpaceIndex + 1), out int number))
                {
                    ws.Cells["D1"].Value = sheetName.Substring(0, lastSpaceIndex).Trim();
                    ws.Cells["H1"].Value = number;
                }
                else { ws.Cells["D1"].Value = sheetName; ws.Cells["H1"].Value = null; }
            }
        }

        public List<RowData> GetSheetDataForPreview(string sheetName)
        {
            if (sheetName == null || !SheetNames.Contains(sheetName)) return new List<RowData>();
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

        public (int, string) RegisterNormalData(string sheetName, Dictionary<string, double?> values, bool isKoryo)
        {
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) throw new Exception($"シート '{sheetName}' に '合計' 行が見つかりません。");
            var (targetRow, insertInfo) = FindTargetRow(ws, totalRowIndex);
            UpdateRowInternal(ws, targetRow, values, isKoryo);
            _dataCache.Remove(sheetName);
            return (targetRow, insertInfo);
        }

        public void UpdateNormalData(string sheetName, int rowIndex, Dictionary<string, double?> values, bool isKoryo)
        {
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            UpdateRowInternal(ws, rowIndex, values, isKoryo);
            _dataCache.Remove(sheetName);
        }

        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r)) { ws.DeleteRow(rowIndex); }
            _dataCache.Remove(sheetName);
        }

        private void UpdateRowInternal(ExcelWorksheet ws, int rowIndex, Dictionary<string, double?> values, bool isKoryo)
        {
            var map = _columnMap.NormalSheet;
            bool isOotsuki = ws.Name.Contains("大月");
            double? yuryoVal = values.GetValueOrDefault("有料キロ(D)");
            int hansoVal = (yuryoVal.HasValue && yuryoVal > 0) ? 1 : 0;
            ws.Cells[rowIndex, map.Day].Value = values.GetValueOrDefault("日(B)");
            ws.Cells[rowIndex, map.HansoCount].Value = hansoVal;
            ws.Cells[rowIndex, map.YuryoKm].Value = yuryoVal;
            ws.Cells[rowIndex, map.MuryoKm].Value = values.GetValueOrDefault("無料キロ(E)");
            ws.Cells[rowIndex, map.IsKoryo].Value = isKoryo ? 1 : (object)null;
            if (isOotsuki) { ws.Cells[rowIndex, map.ShinyaFee].Value = values.GetValueOrDefault("深夜料金(H)"); ws.Cells[rowIndex, map.ShinyaMinutes].Value = null; }
            else { ws.Cells[rowIndex, map.ShinyaFee].Value = null; ws.Cells[rowIndex, map.ShinyaMinutes].Value = values.GetValueOrDefault("深夜時間(K)"); }
        }

        public void RegisterEastData(string sheetName, Dictionary<string, double?> values)
        {
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.EastSheet;
            ws.Cells[map.Jitsudo].Value = values.GetValueOrDefault("延実働車輌数");
            ws.Cells[map.Hanso].Value = values.GetValueOrDefault("搬送回数");
            ws.Cells[map.YuryoKm].Value = values.GetValueOrDefault("有料キロ数");
            ws.Cells[map.MuryoKm].Value = values.GetValueOrDefault("無料キロ数");
            ws.Cells[map.UnsoJisseki].Value = values.GetValueOrDefault("運輸実績");
        }

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
                    ws.Cells[eastMap.Jitsudo].Value = null; ws.Cells[eastMap.Hanso].Value = null; ws.Cells[eastMap.YuryoKm].Value = null;
                    ws.Cells[eastMap.MuryoKm].Value = null; ws.Cells[eastMap.UnsoJisseki].Value = null;
                    logMessages.Add($"[{ws.Name}] のデータをクリアしました。");
                }
            }
            _dataCache.Clear();
            return logMessages;
        }

        public bool CheckRemainingData()
        {
            var map = _columnMap.NormalSheet;
            foreach (var ws in _inputPackage.Workbook.Worksheets)
            {
                if ((ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH")) && ws.Cells[3, map.Day].Value != null) return true;
            }
            return false;
        }

        private static int FindTotalRow(ExcelWorksheet ws)
        {
            if (ws?.Dimension == null) return -1;
            for (int row = ws.Dimension.End.Row; row >= 3; row--) { if (ws.Cells[row, 1].Value?.ToString()?.Contains("合計") == true) return row; }
            return -1;
        }

        private (int targetRow, string insertInfo) FindTargetRow(ExcelWorksheet ws, int totalRowIndex)
        {
            var map = _columnMap.NormalSheet;
            for (int rowNum = 3; rowNum < totalRowIndex; rowNum++) { if (ws.Cells[rowNum, map.Day].Value == null) return (rowNum, ""); }
            ws.InsertRow(totalRowIndex, 1);
            return (totalRowIndex, "空き行がないため、合計行の上に新しい行を挿入します。");
        }

        private static int? GetNullableInt(object val) => val == null ? null : (int?)Convert.ToDouble(val);
        private static double? GetNullableDouble(object val) => val == null ? null : Convert.ToDouble(val);
    }
}