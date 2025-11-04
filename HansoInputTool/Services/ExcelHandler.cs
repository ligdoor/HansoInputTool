using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HansoInputTool.Models;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    public class ExcelHandler
    {
        private readonly string _filePath;
        private readonly ColumnMapping _columnMap;
        private ExcelPackage _excelPackage;
        private readonly Dictionary<string, List<RowData>> _dataCache = new();

        public List<string> SheetNames { get; private set; }

        public ExcelHandler(string filePath, ColumnMapping columnMap)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _filePath = filePath;
            _columnMap = columnMap; // マッピング情報を保持
            Load();
        }

        public void Load()
        {
            _excelPackage?.Dispose();
            var fileInfo = new FileInfo(_filePath);
            _excelPackage = new ExcelPackage(fileInfo);
            SheetNames = _excelPackage.Workbook.Worksheets.Select(ws => ws.Name).ToList();
            _dataCache.Clear();
        }

        public void Save()
        {
            _excelPackage.Save();
        }

        public List<RowData> GetSheetDataForPreview(string sheetName)
        {
            if (_dataCache.ContainsKey(sheetName)) return _dataCache[sheetName];
            if (sheetName == null || !SheetNames.Contains(sheetName)) return new List<RowData>();

            var ws = _excelPackage.Workbook.Worksheets[sheetName];
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
                    D_YuryoKm = GetNullableInt(ws.Cells[rowIndex, map.YuryoKm].Value), // intとして読み込む
                    E_MuryoKm = GetNullableInt(ws.Cells[rowIndex, map.MuryoKm].Value), // intとして読み込む
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

        public (int targetRow, string insertInfo) RegisterNormalData(string sheetName, Dictionary<string, double?> values, bool isKoryo)
        {
            var ws = _excelPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) throw new Exception($"シート '{sheetName}' に '合計' 行が見つかりません。");

            var (targetRow, insertInfo) = FindTargetRow(ws, totalRowIndex);
            UpdateRowInternal(ws, targetRow, values, isKoryo);
            _dataCache.Remove(sheetName);
            return (targetRow, insertInfo);
        }

        public void UpdateNormalData(string sheetName, int rowIndex, Dictionary<string, double?> values, bool isKoryo)
        {
            var ws = _excelPackage.Workbook.Worksheets[sheetName];
            UpdateRowInternal(ws, rowIndex, values, isKoryo);
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
        }

        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            var ws = _excelPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r))
            {
                ws.DeleteRow(rowIndex);
            }
            _dataCache.Remove(sheetName);
        }

        public void RegisterEastData(string sheetName, Dictionary<string, double?> values)
        {
            var ws = _excelPackage.Workbook.Worksheets[sheetName];
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

            foreach (var ws in _excelPackage.Workbook.Worksheets)
            {
                if (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車"))
                {
                    var totalRowIndex = FindTotalRow(ws);
                    if (totalRowIndex != -1)
                    {
                        for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
                        {
                            foreach (int colIndex in new[] { normalMap.Day, normalMap.HansoCount, normalMap.YuryoKm, normalMap.MuryoKm, normalMap.ShinyaFee, normalMap.ShinyaMinutes, normalMap.IsKoryo })
                            {
                                ws.Cells[rowIndex, colIndex].Value = null;
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

        public bool CheckRemainingData()
        {
            var map = _columnMap.NormalSheet;
            foreach (var ws in _excelPackage.Workbook.Worksheets)
            {
                if ((ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車")) && ws.Cells[3, map.Day].Value != null)
                {
                    return true;
                }
            }
            return false;
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

        private (int targetRow, string insertInfo) FindTargetRow(ExcelWorksheet ws, int totalRowIndex)
        {
            var map = _columnMap.NormalSheet;
            for (int rowNum = 3; rowNum < totalRowIndex; rowNum++)
            {
                if (ws.Cells[rowNum, map.Day].Value == null)
                {
                    return (rowNum, "");
                }
            }
            ws.InsertRow(totalRowIndex, 1);
            return (totalRowIndex, "空き行がないため、合計行の上に新しい行を挿入します。");
        }

        private static int? GetNullableInt(object val) => val == null ? null : Convert.ToInt32(Convert.ToDouble(val)); // doubleを経由して安全にintへ
        private static double? GetNullableDouble(object val) => val == null ? null : Convert.ToDouble(val);
    }
}