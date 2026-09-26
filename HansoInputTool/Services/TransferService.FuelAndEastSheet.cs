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
    /// TransferService の分割定義：給油管理シートへの書き込みと、東日本セレモニー系シートの転記。
    /// </summary>
    public partial class TransferService
    {
        // ===== 給油管理表への書き込み =====
        // Excelのシートタブ名は表内のタイトル文字列と異なる場合があるため、
        // 「給油管理」という文字列を含むシートを対象として検索する（例:「給油管理表」「寝台車　給油管理」など）。
        private const string FuelSheetKeyword = "給油管理";

        /// <summary>
        /// DBに登録された給油記録を、実績月報（wbGeppo）内の給油管理シートへ書き込む。
        /// 車両ごとの列位置はハードコードせず、車両名の見出しセルを起点に実行時に探す
        /// （Template.xlsxの列レイアウトが変わっても追従できるようにするため）。
        /// </summary>
        private void WriteFuelSheet(ExcelPackage wbGeppo, List<FuelRecord> fuelRecords)
        {
            if (fuelRecords == null || fuelRecords.Count == 0) return;

            var ws = wbGeppo.Workbook.Worksheets.FirstOrDefault(s => s.Name.Contains(FuelSheetKeyword));
            if (ws == null)
            {
                Logger.Warn($"「{FuelSheetKeyword}」を含むシートが見つからないため、給油記録の書き込みをスキップしました。");
                return;
            }

            foreach (var group in fuelRecords.GroupBy(f => f.VehicleSheetName))
            {
                var block = FindFuelVehicleBlock(ws, group.Key);
                if (block == null)
                {
                    Logger.Warn($"「{FuelSheetKeyword}」を含むシートに車両「{group.Key}」の列が見つからないため、この車両の給油記録はスキップしました。");
                    continue;
                }

                var (dayCol, kmCol, litersCol, firstDataRow) = block.Value;

                int row = firstDataRow;
                foreach (var fuel in group.OrderBy(f => f.Day))
                {
                    // 既に値が入っている行はスキップして次の空き行へ
                    while (ws.Cells[row, dayCol].Value != null) row++;
                    ws.Cells[row, dayCol].Value     = fuel.Day;
                    ws.Cells[row, kmCol].Value      = fuel.OdometerKm;
                    ws.Cells[row, litersCol].Value  = fuel.Liters;
                    row++;
                }
                Logger.Info($"「{FuelSheetKeyword}」を含むシート[{group.Key}] に給油記録{group.Count()}件を書き込みました。");
            }
        }

        /// <summary>
        /// 給油管理表シート内で指定車両の列ブロック（日・給油時Km・給油㍑数の列番号とデータ開始行）を探す。
        /// 車両名の見出しセルを見つけ、その付近で「日」ヘッダーの位置から列を特定する。
        /// </summary>
        private (int dayCol, int kmCol, int litersCol, int firstDataRow)? FindFuelVehicleBlock(ExcelWorksheet ws, string vehicleSheetName)
        {
            if (ws.Dimension == null) return null;

            for (int row = 1; row <= Math.Min(10, ws.Dimension.End.Row); row++)
            {
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    var text = ws.Cells[row, col].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // Excel側の見出し（例:"寝台車 29"）がDB側の車両シート名の末尾または一部と一致するか
                    if (!vehicleSheetName.EndsWith(text) && !vehicleSheetName.Contains(text)) continue;

                    // 見出しの下、数行以内で「日」ヘッダーの列を探す
                    for (int r2 = row + 1; r2 <= Math.Min(row + 4, ws.Dimension.End.Row); r2++)
                    {
                        for (int c2 = col; c2 <= Math.Min(col + 4, ws.Dimension.End.Column); c2++)
                        {
                            if (ws.Cells[r2, c2].Value?.ToString() == "日")
                                return (c2, c2 + 1, c2 + 2, r2 + 1);
                        }
                    }
                }
            }
            return null;
        }


        private void ProcessEastSheet(ExcelPackage wbInput, ExcelPackage wbShukei, string sheetName, ColumnMapping columnMap)
        {
            // [廃車済み車両対応] 該当シートが集計ファイルに無ければ自動生成してから書き込む
            var wsShukei = EnsureShukeiSheet(wbShukei, sheetName);
            if (wsShukei == null) return;
            var wsIn = wbInput.Workbook.Worksheets[sheetName];
            var shukeiMap = columnMap.ShukeiSheet;
            wsShukei.Cells[shukeiMap.Days].Value = wsIn.Cells[shukeiMap.Days].Value;
            wsShukei.Cells[shukeiMap.Hanso].Value = wsIn.Cells[shukeiMap.Hanso].Value;
            wsShukei.Cells[shukeiMap.YuryoKm].Value = wsIn.Cells[shukeiMap.YuryoKm].Value;
            wsShukei.Cells[shukeiMap.MuryoKm].Value = wsIn.Cells[shukeiMap.MuryoKm].Value;
            wsShukei.Cells[shukeiMap.Total].Value = wsIn.Cells[shukeiMap.Total].Value;
            Logger.Info($"[{sheetName}] の値を転記しました。");
        }
    }
}
