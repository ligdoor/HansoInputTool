using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// Excelファイルの読み書き・シート管理を担当するクラス。
    /// partial classで3ファイルに分割:
    ///   ExcelHandler.cs             - 基本操作・データ読み書き（このファイル）
    ///   ExcelHandler.SheetSync.cs   - シート同期・並べ替え・命名
    ///   ExcelHandler.MonthlySummary.cs - 月間集計シート更新
    /// </summary>
    public partial class ExcelHandler : IDisposable
    {
        private bool _disposed = false;

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
            _inputFilePath    = inputFilePath;
            _templateFilePath = templateFilePath;
            _columnMap        = columnMap;
            Load();
        }

        #region 基本操作

        public void Load()
        {
            _inputPackage?.Dispose();
            _templatePackage?.Dispose();
            _inputPackage    = new ExcelPackage(new FileInfo(_inputFilePath));
            _templatePackage = new ExcelPackage(new FileInfo(_templateFilePath));
            _dataCache.Clear();
        }

        public void Save()
        {
            _inputPackage?.Save();
            _templatePackage?.Save();
        }

        public bool TemplateSheetExists(string sheetName)
            => _templatePackage.Workbook.Worksheets.Any(ws => ws.Name == sheetName);

        public List<string> SheetNames => _inputPackage?.Workbook.Worksheets
            .Where(ws => !ws.Name.Contains("登録") && !IsTemplateSheet(ws.Name))
            .Select(ws => ws.Name)
            .ToList() ?? new List<string>();

        public List<string> GetVehicleSheetNames()
            => _inputPackage.Workbook.Worksheets
                .Where(s => !s.Name.Contains("登録") && !IsTemplateSheet(s.Name))
                .Select(s => s.Name)
                .ToList();

        #endregion

        #region データ読み取り（プレビュー用）

        public List<RowData> GetSheetDataForPreview(string sheetName)
        {
            if (sheetName == null || !_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                return new List<RowData>();
            if (_dataCache.ContainsKey(sheetName)) return _dataCache[sheetName];

            var ws            = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return new List<RowData>();

            var data       = new List<RowData>();
            var map        = _columnMap.NormalSheet;
            bool isOotsuki = sheetName.Contains("大月");

            for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
            {
                if (ws.Cells[rowIndex, map.Day].Value == null && ws.Cells[rowIndex, map.YuryoKm].Value == null) continue;

                var rowData = new RowData
                {
                    RowIndex         = rowIndex,
                    B_Day            = GetNullableInt(ws.Cells[rowIndex, map.Day].Value),
                    C_Hanso          = GetNullableInt(ws.Cells[rowIndex, map.HansoCount].Value),
                    D_YuryoKm        = GetNullableInt(ws.Cells[rowIndex, map.YuryoKm].Value),
                    E_MuryoKm        = GetNullableInt(ws.Cells[rowIndex, map.MuryoKm].Value),
                    H_LateFeeOotsuki = GetNullableInt(ws.Cells[rowIndex, map.ShinyaFee].Value),
                    K_LateMinutes    = GetNullableInt(ws.Cells[rowIndex, map.ShinyaMinutes].Value),
                    L_IsKoryo        = GetNullableInt(ws.Cells[rowIndex, map.IsKoryo].Value),
                    M_IsEmbalming    = GetNullableInt(ws.Cells[rowIndex, map.IsEmbalming].Value)
                };
                rowData.LateValueText = isOotsuki
                    ? rowData.H_LateFeeOotsuki?.ToString()
                    : rowData.K_LateMinutes?.ToString();
                data.Add(rowData);
            }

            _dataCache[sheetName] = data;
            return data;
        }

        #endregion

        #region データ書き込み

        public (int targetRow, string insertInfo) RegisterNormalData(string sheetName, Dictionary<string, double?> values, bool isKoryo, bool isEmbalming)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");

            var ws            = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) throw new Exception($"シート '{sheetName}' に '合計' 行が見つかりません。");

            var map = _columnMap.NormalSheet;
            int targetRow = -1;
            for (int r = 3; r < totalRowIndex; r++)
                if (ws.Cells[r, map.Day].Value == null) { targetRow = r; break; }

            string insertInfo = "";
            if (targetRow == -1)
            {
                ws.InsertRow(totalRowIndex, 1);
                targetRow  = totalRowIndex;
                insertInfo = "空き行がないため、合計行の上に新しい行を挿入します。";
            }

            WriteNormalValues(ws, targetRow, map, values, isKoryo, isEmbalming, sheetName.Contains("大月"));
            InvalidateCache(sheetName);
            return (targetRow, insertInfo);
        }

        public void UpdateNormalData(string sheetName, int rowIndex, Dictionary<string, double?> values, bool isKoryo, bool isEmbalming)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            WriteNormalValues(_inputPackage.Workbook.Worksheets[sheetName], rowIndex, _columnMap.NormalSheet, values, isKoryo, isEmbalming, sheetName.Contains("大月"));
            InvalidateCache(sheetName);
        }

        public void RegisterEastData(string sheetName, Dictionary<string, double?> values)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws  = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.EastSheet;
            ws.Cells[map.Jitsudo].Value     = values.GetValueOrDefault("延実働車輌数");
            ws.Cells[map.Hanso].Value       = values.GetValueOrDefault("搬送回数");
            ws.Cells[map.YuryoKm].Value     = values.GetValueOrDefault("有料キロ数");
            ws.Cells[map.MuryoKm].Value     = values.GetValueOrDefault("無料キロ数");
            ws.Cells[map.UnsoJisseki].Value = values.GetValueOrDefault("運輸実績");
        }

        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r))
                ws.DeleteRow(rowIndex);
            InvalidateCache(sheetName);
        }

        public List<string> ClearData()
        {
            var logMessages = new List<string>();
            var normalMap   = _columnMap.NormalSheet;
            var eastMap     = _columnMap.EastSheet;

            foreach (var ws in _inputPackage.Workbook.Worksheets)
            {
                if (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                {
                    var totalRowIndex = FindTotalRow(ws);
                    if (totalRowIndex != -1)
                    {
                        for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
                            foreach (int col in new[] { normalMap.Day, normalMap.HansoCount, normalMap.YuryoKm, normalMap.MuryoKm, normalMap.ShinyaFee, normalMap.ShinyaMinutes, normalMap.IsKoryo, normalMap.IsEmbalming })
                                if (col > 0) ws.Cells[rowIndex, col].Value = null;
                        logMessages.Add($"[{ws.Name}] の入力値をクリアしました。");
                    }
                }
                else if (ws.Name.Contains("東日本"))
                {
                    ws.Cells[eastMap.Jitsudo].Value     = null;
                    ws.Cells[eastMap.Hanso].Value       = null;
                    ws.Cells[eastMap.YuryoKm].Value     = null;
                    ws.Cells[eastMap.MuryoKm].Value     = null;
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
            foreach (var ws in _inputPackage.Workbook.Worksheets)
                if ((ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                    && ws.Cells[3, map.Day].Value != null)
                    return true;
            return false;
        }

        #endregion

        #region 内部ヘルパー（他のpartialファイルからも使用）

        private void WriteNormalValues(ExcelWorksheet ws, int row, SheetColumnMap map,
            Dictionary<string, double?> values, bool isKoryo, bool isEmbalming, bool isOotsuki)
        {
            double? yuryoVal = values.GetValueOrDefault("有料キロ(D)");
            int hansoVal     = (yuryoVal.HasValue && yuryoVal > 0) ? 1 : 0;

            ws.Cells[row, map.Day].Value        = values.GetValueOrDefault("日(B)");
            ws.Cells[row, map.HansoCount].Value = hansoVal;
            ws.Cells[row, map.YuryoKm].Value    = yuryoVal;
            ws.Cells[row, map.MuryoKm].Value    = values.GetValueOrDefault("無料キロ(E)");
            ws.Cells[row, map.IsKoryo].Value     = isKoryo ? 1 : (object)null;
            ws.Cells[row, map.IsEmbalming].Value = isEmbalming ? 1 : (object)null;

            if (isOotsuki)
            {
                ws.Cells[row, map.ShinyaFee].Value     = values.GetValueOrDefault("深夜料金(H)");
                ws.Cells[row, map.ShinyaMinutes].Value = null;
            }
            else
            {
                ws.Cells[row, map.ShinyaFee].Value     = null;
                ws.Cells[row, map.ShinyaMinutes].Value = values.GetValueOrDefault("深夜時間(K)");
            }
        }

        private void InvalidateCache(string sheetName)
        {
            if (_dataCache.ContainsKey(sheetName)) _dataCache.Remove(sheetName);
        }

        /// <summary>
        /// 指定シートのエンバーミング件数（M列=1 の行数）を返す
        /// </summary>
        public int GetEmbalmingCount(string sheetName)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) return 0;
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return 0;
            int count = 0;
            int col = _columnMap.NormalSheet.IsEmbalming;
            for (int r = 3; r < totalRowIndex; r++)
                if (GetNullableInt(ws.Cells[r, col].Value) == 1) count++;
            return count;
        }

        internal static int FindTotalRow(ExcelWorksheet ws)
        {
            if (ws?.Dimension == null) return -1;
            for (int row = ws.Dimension.End.Row; row >= 3; row--)
                if (ws.Cells[row, 1].Value?.ToString()?.Contains("合計") == true) return row;
            return -1;
        }

        internal static int? GetNullableInt(object val)
        {
            if (val == null) return null;
            if (val is int i)     return i;
            if (val is long l)    return (int)l;
            if (val is double d)  return (int)d;
            if (val is decimal m) return (int)m;

            var s = val.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            s = s.Replace(",", "").Replace("，", "");

            if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double parsed) ||
                double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed))
                return (int)parsed;

            Logger.Warn($"GetNullableInt: 非数値フィールドをパースできませんでした: '{s}'");
            return null;
        }

        internal static double? GetNullableDouble(object val)
            => val == null ? null : Convert.ToDouble(val);

        internal static bool IsTemplateSheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (sheetName.Replace(" ", "").StartsWith("template", StringComparison.OrdinalIgnoreCase)) return true;
            if (sheetName.IndexOf("テンプレート", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        internal static bool NeedQuotes(string sheetName)
            => sheetName.Contains(" ") || sheetName.Contains("-") || sheetName.Contains("(")
            || sheetName.Contains(")") || sheetName.Contains("'") || sheetName.Contains("!")
            || sheetName.Contains("#");

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                _inputPackage?.Dispose();
                _templatePackage?.Dispose();
                _disposed = true;
            }
        }
    }
}
