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

        public List<string> GetVehicleSheetNames()
        {
            return SheetNames.Where(s => !s.Contains("登録")).ToList();
        }

        public void AddVehicleSheet(string newSheetName, string templateSheetName)
        {
            Logger.Info($"シート追加処理を開始: {newSheetName} (テンプレート: {templateSheetName})");

            var inputTemplateWs = _inputPackage.Workbook.Worksheets[templateSheetName];
            if (inputTemplateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateSheetName}' がInput.xlsxに見つかりません。");
            int inputIndex = GetInsertIndex(_inputPackage, templateSheetName);
            var newInputWs = _inputPackage.Workbook.Worksheets.Copy(inputTemplateWs.Name, newSheetName);
            _inputPackage.Workbook.Worksheets.MoveAfter(newInputWs.Index, inputIndex);
            UpdateSheetCells(newInputWs);

            var templateTemplateWs = _templatePackage.Workbook.Worksheets[templateSheetName];
            if (templateTemplateWs == null) throw new FileNotFoundException($"コピー元のシート '{templateSheetName}' がTemplate.xlsxに見つかりません。");
            int templateIndex = GetInsertIndex(_templatePackage, templateSheetName);
            var newTemplateWs = _templatePackage.Workbook.Worksheets.Copy(templateTemplateWs.Name, newSheetName);
            _templatePackage.Workbook.Worksheets.MoveAfter(newTemplateWs.Index, templateIndex);
        }

        public void DeleteVehicleSheet(string sheetName)
        {
            Logger.Info($"シート削除処理を開始: {sheetName}");
            var inputWs = _inputPackage.Workbook.Worksheets[sheetName];
            if (inputWs != null) _inputPackage.Workbook.Worksheets.Delete(inputWs);

            var templateWs = _templatePackage.Workbook.Worksheets[sheetName];
            if (templateWs != null) _templatePackage.Workbook.Worksheets.Delete(templateWs);
        }

        public void RenameVehicleSheet(string oldSheetName, string newSheetName)
        {
            Logger.Info($"シート名変更処理を開始: {oldSheetName} -> {newSheetName}");
            var inputWs = _inputPackage.Workbook.Worksheets[oldSheetName];
            if (inputWs != null)
            {
                inputWs.Name = newSheetName;
                UpdateSheetCells(inputWs);
            }

            var templateWs = _templatePackage.Workbook.Worksheets[oldSheetName];
            if (templateWs != null) templateWs.Name = newSheetName;
        }

        public void UpdateMonthlySummarySheet(List<string> updatedSheetNames)
        {
            Logger.Info("月間集計シートの同期処理を開始します。");
            var summarySheet = _templatePackage.Workbook.Worksheets["月間集計"];
            if (summarySheet == null)
            {
                Logger.Warn("Template.xlsxに '月間集計' シートが見つかりませんでした。");
                return;
            }

            var totalRow = FindTotalRow(summarySheet);
            if (totalRow == -1)
            {
                Logger.Warn("'月間集計' シートに合計行が見つかりませんでした。");
                return;
            }

            var currentVehicles = new Dictionary<string, (string Branch, string Number)>();
            for (int row = 4; row < totalRow; row++)
            {
                var branch = summarySheet.Cells[row, 2].Text;
                var number = summarySheet.Cells[row, 3].Text;
                if (!string.IsNullOrEmpty(branch))
                {
                    currentVehicles[BuildSheetName(branch, number)] = (branch, number);
                }
            }

            var sheetsToDelete = currentVehicles.Keys.Except(updatedSheetNames).ToList();
            if (sheetsToDelete.Any())
            {
                for (int row = totalRow - 1; row >= 4; row--)
                {
                    var branch = summarySheet.Cells[row, 2].Text;
                    var number = summarySheet.Cells[row, 3].Text;
                    var sheetName = BuildSheetName(branch, number);
                    if (sheetsToDelete.Contains(sheetName))
                    {
                        summarySheet.DeleteRow(row, 1);
                        Logger.Info($"月間集計シートから行を削除: {sheetName}");
                    }
                }
            }

            var sheetsToAdd = updatedSheetNames.Except(currentVehicles.Keys).ToList();
            foreach (var sheetName in sheetsToAdd)
            {
                var (branch, number) = ParseSheetNameToBranchAndNumber(sheetName);
                string categoryKey = GetCategoryKey(sheetName);

                int lastRowOfCategory = 3;
                for (int row = 4; row < FindTotalRow(summarySheet); row++)
                {
                    var existingBranch = summarySheet.Cells[row, 2].Text;
                    if (GetCategoryKey(existingBranch) == categoryKey)
                    {
                        lastRowOfCategory = row;
                    }
                }

                int insertRow = lastRowOfCategory + 1;
                summarySheet.InsertRow(insertRow, 1, lastRowOfCategory);
                summarySheet.Cells[insertRow, 2].Value = branch;
                summarySheet.Cells[insertRow, 3].Value = int.TryParse(number, out int num) ? num : (object)number;
                Logger.Info($"月間集計シートに行を挿入: {sheetName} at row {insertRow}");
            }

            totalRow = FindTotalRow(summarySheet);
            for (int row = 4; row < totalRow; row++)
            {
                summarySheet.Cells[row, 1].Value = $"No.{row - 3}";
            }
        }

        private (string Branch, string Number) ParseSheetNameToBranchAndNumber(string sheetName)
        {
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("東日本", numberMatch.Success ? numberMatch.Value : "");
            }
            if (sheetName.Contains("CH"))
            {
                var lastSpaceIndex = sheetName.LastIndexOf(' ');
                if (lastSpaceIndex > -1 && int.TryParse(sheetName.Substring(lastSpaceIndex + 1), out _))
                {
                    return (sheetName.Substring(0, lastSpaceIndex), sheetName.Substring(lastSpaceIndex + 1));
                }
            }
            var parts = sheetName.Split(' ');
            if (parts.Length > 1 && int.TryParse(parts.Last(), out _))
            {
                return (string.Join(" ", parts.Take(parts.Length - 1)), parts.Last());
            }
            return (sheetName, "");
        }

        private string BuildSheetName(string branch, string number)
        {
            if (branch == "東日本") { return $"東日本セレモニー {number}"; }
            if (!string.IsNullOrEmpty(number)) { return $"{branch} {number}"; }
            return branch;
        }

        private int GetInsertIndex(ExcelPackage package, string baseSheetName)
        {
            string categoryKey = GetCategoryKey(baseSheetName);
            var categorySheets = package.Workbook.Worksheets.Where(ws => GetCategoryKey(ws.Name) == categoryKey).ToList();
            if (categorySheets.Any()) { return categorySheets.Max(ws => ws.Index); }
            return package.Workbook.Worksheets.Count;
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
                else
                {
                    ws.Cells["D1"].Value = sheetName;
                    ws.Cells["H1"].Value = null;
                }
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

        // ↓↓↓ 不足していた DeleteRows メソッドを復活 ↓↓↓
        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r))
            {
                ws.DeleteRow(rowIndex);
            }
            _dataCache.Remove(sheetName);
        }
        // ↑↑↑ ここまでが復活させたメソッド ↑↑↑

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