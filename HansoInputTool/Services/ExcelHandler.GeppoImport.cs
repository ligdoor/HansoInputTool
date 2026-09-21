using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// ExcelHandler の partial クラス：実績月報ファイルの読込（「月報読込」ボタン）
    /// </summary>
    public partial class ExcelHandler
    {
        /// <summary>
        /// [給油管理表消失バグ対策]
        /// 以前の実装は、選択した実績月報ファイルで Input.xlsx を File.Copy によって丸ごと置き換えていた。
        /// しかし実績月報ファイルには、給油管理表・Template1/Template2（ひな形）・月間集計といった
        /// Input.xlsx にしか存在しない作業用シートが含まれていないため、丸ごと置き換えるとこれらの
        /// シートが Input.xlsx から消えてしまっていた。
        ///
        /// また、EPPlus には「テーブル(ExcelTable)を含むシートを別ファイルへシート単位でコピーすると
        /// ファイルが壊れる」既知の不具合があるため（VehicleAnnualSummaryService.cs 参照）、
        /// シートそのものをコピーするのではなく、同名シートが両方に存在する場合にセルの値だけを
        /// 1件ずつ転記する方式にしている。
        ///
        /// 対象は通常系（寝台車・霊柩車・CH）・東日本セレモニーシートのみ。
        /// 給油管理表・Template1/Template2・月間集計・登録シートは一切変更しない。
        /// 実績月報側にのみ存在するシート（Input.xlsx側に同名シートが無い場合）は取り込まない。
        /// </summary>
        public List<string> ImportFromGeppoFile(string sourceFilePath, out List<Models.FuelRecord> importedFuelRecords)
        {
            var logMessages = new List<string>();
            var normalMap   = _columnMap.NormalSheet;
            var eastMap     = _columnMap.EastSheet;
            var flags       = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();
            var flagCols    = flags.Select(f => f.ExcelColumn).Where(c => c > 0).Distinct().ToArray();

            using var sourcePackage = new ExcelPackage(new FileInfo(sourceFilePath));

            foreach (var destWs in _inputPackage.Workbook.Worksheets)
            {
                if (IsProtectedSheet(destWs.Name) || destWs.Name.Contains("登録")) continue;

                var srcWs = sourcePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name == destWs.Name);
                if (srcWs == null) continue; // 実績月報側に同名シートが無ければInput.xlsx側はそのまま維持する

                if (destWs.Name.Contains("寝台車") || destWs.Name.Contains("霊柩車") || destWs.Name.Contains("CH"))
                {
                    var destTotalRow = FindTotalRow(destWs);
                    var srcTotalRow  = FindTotalRow(srcWs);
                    if (destTotalRow == -1 || srcTotalRow == -1) continue;

                    var cols = new[] {
                        normalMap.Day, normalMap.HansoCount, normalMap.YuryoKm, normalMap.MuryoKm,
                        normalMap.KihonFee, normalMap.SokoFee, normalMap.ShinyaFee, normalMap.TotalFee,
                        normalMap.ShinyaMinutes
                    }.Concat(flagCols).Where(c => c > 0).Distinct().ToArray();

                    int maxRow = Math.Max(destTotalRow, srcTotalRow);
                    for (int row = 3; row <= maxRow; row++)
                        foreach (int col in cols)
                            destWs.Cells[row, col].Value = row <= srcTotalRow ? srcWs.Cells[row, col].Value : null;

                    logMessages.Add($"[{destWs.Name}] のデータを実績月報から取り込みました。");
                }
                else if (destWs.Name.Contains("東日本"))
                {
                    CopyCellIfAddressSet(srcWs, destWs, eastMap.Jitsudo);
                    CopyCellIfAddressSet(srcWs, destWs, eastMap.Hanso);
                    CopyCellIfAddressSet(srcWs, destWs, eastMap.YuryoKm);
                    CopyCellIfAddressSet(srcWs, destWs, eastMap.MuryoKm);
                    CopyCellIfAddressSet(srcWs, destWs, eastMap.UnsoJisseki);
                    logMessages.Add($"[{destWs.Name}] のデータを実績月報から取り込みました。");
                }
            }

            // [給油情報の復元] 実績月報側の給油管理表に記録が残っていれば読み取っておく。
            // ここでは読み取るだけで、DB/Excelどちらへ書き戻すかは呼び出し元（LoadGeppoFile）に任せる
            // （DB使用時はセッション作成後でないとDBへ挿入できないため）。
            importedFuelRecords = ReadFuelRecordsFromSource(sourcePackage);
            if (importedFuelRecords.Count > 0)
                logMessages.Add($"「{FuelSheetKeyword}」を含むシートから給油記録{importedFuelRecords.Count}件を読み取りました。");

            _dataCache.Clear();
            return logMessages;
        }

        private static void CopyCellIfAddressSet(ExcelWorksheet src, ExcelWorksheet dest, string cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress)) return;
            dest.Cells[cellAddress].Value = src.Cells[cellAddress].Value;
        }

        // Excelのシートタブ名は表内のタイトル文字列と異なる場合があるため、
        // 「給油管理」という文字列を含むシートを対象として検索する（TransferService.csと同じ規約）。
        private const string FuelSheetKeyword = "給油管理";

        /// <summary>
        /// [給油情報の復元] 実績月報（sourcePackage）内の給油管理表シートから、給油記録（日・Km・㍑）を読み取る。
        /// 車両ごとの列位置はハードコードせず、車両名の見出しセルを起点に実行時に探す
        /// （TransferService.WriteFuelSheet/FindFuelVehicleBlockと同じ考え方）。
        /// 見出しの短縮名（例:"寝台車 29"）は、Input.xlsx側の実際のシート名（例:"CH富士吉田 寝台車 29"）と
        /// 末尾一致・部分一致で対応付ける。
        /// </summary>
        private List<Models.FuelRecord> ReadFuelRecordsFromSource(ExcelPackage sourcePackage)
        {
            var result = new List<Models.FuelRecord>();

            var srcWs = sourcePackage.Workbook.Worksheets.FirstOrDefault(s => s.Name.Contains(FuelSheetKeyword));
            if (srcWs == null || srcWs.Dimension == null) return result;

            var vehicleSheetNames = _inputPackage.Workbook.Worksheets
                .Where(ws => !IsProtectedSheet(ws.Name) && !ws.Name.Contains("登録"))
                .Select(ws => ws.Name)
                .ToList();

            for (int row = 1; row <= Math.Min(10, srcWs.Dimension.End.Row); row++)
            {
                for (int col = 1; col <= srcWs.Dimension.End.Column; col++)
                {
                    var headingText = srcWs.Cells[row, col].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(headingText)) continue;

                    var vehicleSheetName = vehicleSheetNames.FirstOrDefault(
                        n => n.EndsWith(headingText) || n.Contains(headingText));
                    if (vehicleSheetName == null) continue;

                    for (int r2 = row + 1; r2 <= Math.Min(row + 4, srcWs.Dimension.End.Row); r2++)
                    {
                        bool found = false;
                        for (int c2 = col; c2 <= Math.Min(col + 4, srcWs.Dimension.End.Column); c2++)
                        {
                            if (srcWs.Cells[r2, c2].Value?.ToString() != "日") continue;

                            int dayCol = c2, kmCol = c2 + 1, litersCol = c2 + 2, dataRow = r2 + 1;
                            while (srcWs.Cells[dataRow, dayCol].Value != null)
                            {
                                var dayVal = GetNullableDouble(srcWs.Cells[dataRow, dayCol].Value);
                                if (dayVal.HasValue)
                                {
                                    result.Add(new Models.FuelRecord
                                    {
                                        VehicleSheetName = vehicleSheetName,
                                        Day              = (int)dayVal.Value,
                                        OdometerKm       = GetNullableDouble(srcWs.Cells[dataRow, kmCol].Value) ?? 0,
                                        Liters           = GetNullableDouble(srcWs.Cells[dataRow, litersCol].Value) ?? 0
                                    });
                                }
                                dataRow++;
                            }
                            found = true;
                            break;
                        }
                        if (found) break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// [給油情報の復元] Excel直接入力モード用：読み取った給油記録をInput.xlsx自身の給油管理表シートへ
        /// 書き込む（DB使用時はDBへ挿入するため、こちらは呼ばない）。
        /// </summary>
        public void WriteFuelRecordsToInputSheet(List<Models.FuelRecord> fuelRecords)
        {
            if (fuelRecords == null || fuelRecords.Count == 0) return;

            var ws = _inputPackage.Workbook.Worksheets.FirstOrDefault(s => s.Name.Contains(FuelSheetKeyword));
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
                    while (ws.Cells[row, dayCol].Value != null) row++; // 既に値がある行はスキップ
                    ws.Cells[row, dayCol].Value    = fuel.Day;
                    ws.Cells[row, kmCol].Value     = fuel.OdometerKm;
                    ws.Cells[row, litersCol].Value = fuel.Liters;
                    row++;
                }
                Logger.Info($"Input.xlsx「{FuelSheetKeyword}」を含むシート[{group.Key}] に給油記録{group.Count()}件を書き込みました。");
            }
        }

        /// <summary>
        /// TransferService.FindFuelVehicleBlockと同じアルゴリズム（給油管理表シート内で指定車両の
        /// 列ブロック：日・給油時Km・給油㍑数の列番号とデータ開始行を探す）。
        /// </summary>
        private static (int dayCol, int kmCol, int litersCol, int firstDataRow)? FindFuelVehicleBlock(
            ExcelWorksheet ws, string vehicleSheetName)
        {
            if (ws.Dimension == null) return null;

            for (int row = 1; row <= Math.Min(10, ws.Dimension.End.Row); row++)
            {
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    var text = ws.Cells[row, col].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    if (!vehicleSheetName.EndsWith(text) && !vehicleSheetName.Contains(text)) continue;

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
    }
}
