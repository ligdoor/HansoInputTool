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
    public class TransferProgressReport
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public string Message { get; set; }
    }

    public class TransferService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public async Task ExecuteAsync(
            string workInputFile,
            string bundledTemplateFile,
            string outputDir,
            int period,
            int month,
            int rNum,
            List<string> allSheetNames,
            Dictionary<string, RateInfo> rates,
            ColumnMapping columnMap,
            IProgress<TransferProgressReport> progress,
            FlagDefinitionService flagService = null,
            DatabaseService dbService = null)
        {
            await Task.Run(() =>
            {
                string folderName = $"{period}期 {month}月 R{rNum} アルス搬送・霊柩車　実績月報";
                string finalOutputDir = Path.Combine(outputDir, folderName);
                Directory.CreateDirectory(finalOutputDir);

                string geppoFilename = $"{period}期 {month}月 R{rNum} アルス搬送・霊柩車　実績月報.xlsx";
                string geppoFilepath = Path.Combine(finalOutputDir, geppoFilename);
                File.Copy(workInputFile, geppoFilepath, true);
                Logger.Info($"実績月報ファイルをコピーしました: {geppoFilepath}");

                string shukeiFilename = $"{period}期 {month}月 R{rNum} アルス搬送・霊柩車　実績月報集計.xlsx";
                string shukeiFilepath = Path.Combine(finalOutputDir, shukeiFilename);
                File.Copy(bundledTemplateFile, shukeiFilepath, true);
                Logger.Info($"集計ファイルをコピーしました: {shukeiFilepath}");

                using var wbInput = new ExcelPackage(new FileInfo(workInputFile));
                using var wbGeppo = new ExcelPackage(new FileInfo(geppoFilepath));
                using var wbShukei = new ExcelPackage(new FileInfo(shukeiFilepath));

                var sheetsToProcess = allSheetNames?.Where(s => !s.Contains("登録")).ToList() ?? new List<string>();
                int totalSheets = sheetsToProcess.Count;
                int processedCount = 0;

                progress.Report(new TransferProgressReport { Current = 0, Total = totalSheets, Message = $"転記処理を開始します（対象: {totalSheets}シート）" });
                Logger.Info($"転記処理を開始: {totalSheets}シート");

                foreach (var sheetName in sheetsToProcess)
                {
                    progress.Report(new TransferProgressReport { Current = processedCount, Total = totalSheets, Message = $"[{processedCount + 1}/{totalSheets}] {sheetName} を処理中..." });

                    if (IsNormalSheet(sheetName))
                    {
                        ProcessNormalSheet(wbInput, wbGeppo, wbShukei, sheetName, rates, columnMap, flagService, dbService);
                    }
                    else if (IsEastSheet(sheetName))
                    {
                        ProcessEastSheet(wbInput, wbShukei, sheetName, columnMap);
                    }
                    else
                    {
                        Logger.Warn($"[{sheetName}] はどのシート種別にも該当しないためスキップしました。");
                    }

                    processedCount++;
                    Logger.Info($"[{sheetName}] 完了 ({processedCount}/{totalSheets})");
                }

                progress.Report(new TransferProgressReport { Current = processedCount, Total = totalSheets, Message = $"全{totalSheets}シートの転記が完了しました。保存中..." });

                // 全ての車両シート（寝台車・霊柩車・CH系）に対して
                // A1=R{rNum}、B1=月 を書き込む（C1=期はファイル名に使用するのみ・セルへの記入は不要）
                foreach (var sheetWs in wbShukei.Workbook.Worksheets)
                {
                    string sn = sheetWs.Name;
                    if (IsNormalSheet(sn) || IsEastSheet(sn))
                    {
                        sheetWs.Cells["A1"].Value = $"R{rNum}";
                        sheetWs.Cells["B1"].Value = month;
                        Logger.Info($"集計ファイル [{sn}] に A1=R{rNum}, B1={month} を書き込みました。");
                    }
                }
                foreach (var sheetWs in wbGeppo.Workbook.Worksheets)
                {
                    string sn = sheetWs.Name;
                    if (IsNormalSheet(sn) || IsEastSheet(sn))
                    {
                        sheetWs.Cells["A1"].Value = $"R{rNum}";
                        sheetWs.Cells["B1"].Value = month;
                        Logger.Info($"月報ファイル [{sn}] に A1=R{rNum}, B1={month} を書き込みました。");
                    }
                }

                wbShukei.Save();
                wbGeppo.Save();

                // DB使用時：転記完了後にDBデータをクリア
                if (dbService != null)
                {
                    dbService.ClearAllData();
                    Logger.Info("転記完了のためDBデータをクリアしました。");
                }
            });
        }

        // ===== シート種別の一元判定ヘルパー =====
        private static readonly string[] NormalSheetKeywords = { "寝台車", "霊柩車", "CH富士吉田", "CH大月", "CH東富士" };
        private static readonly string[] EastSheetKeywords = { "東日本セレモニー", "東日本" };

        private static bool IsNormalSheet(string sheetName) =>
            NormalSheetKeywords.Any(kw => sheetName.Contains(kw));

        private static bool IsEastSheet(string sheetName) =>
            EastSheetKeywords.Any(kw => sheetName.Contains(kw));
        // ==========================================

        private void ProcessNormalSheet(ExcelPackage wbInput, ExcelPackage wbGeppo, ExcelPackage wbShukei, string sheetName, Dictionary<string, RateInfo> rates, ColumnMapping columnMap, FlagDefinitionService flagService = null, DatabaseService dbService = null)
        {
            var wsIn = wbInput.Workbook.Worksheets[sheetName];
            var wsGeppo = wbGeppo.Workbook.Worksheets[sheetName];
            var totalRowIdx = FindTotalRow(wsIn);
            if (totalRowIdx == -1) return;

            var normalMap = columnMap.NormalSheet;
            var shukeiMap = columnMap.ShukeiSheet;

            string rateCategory = sheetName.Contains("霊柩車") ? "霊柩車" : "寝台車";
            if (!rates.TryGetValue(rateCategory, out var ratesForSheet))
            {
                Logger.Warn($"シート '{sheetName}' に対応する料金カテゴリ '{rateCategory}' が見つかりませんでした。");
                return;
            }

            bool isOotsuki = sheetName.Contains("大月");
            double totalKihon = 0, totalSoko = 0, totalShinya = 0, totalSum = 0;

            // 金額ありフラグ（WithAmount）を取得して料金計算に使う
            var withAmountFlags = flagService?.Flags
                .Where(f => f.Type == FlagType.WithAmount)
                .ToList() ?? new List<FlagDefinition>();

            // ── DB使用時：DBからデータを読んでwsGeppoに書き込む ──
            if (dbService != null)
            {
                var flags = flagService?.Flags ?? new System.Collections.ObjectModel.ReadOnlyCollection<FlagDefinition>(new List<FlagDefinition>());
                var dbRows = dbService.GetSheetData(sheetName, flags);
                int writeRow = 3;
                foreach (var dbRow in dbRows)
                {
                    int hansoVal = dbRow.C_Hanso ?? 0;
                    double yuryoKm = dbRow.D_YuryoKm ?? 0;
                    double muryoKm = dbRow.E_MuryoKm ?? 0;
                    double rowKihon = 0, rowSoko = 0, rowShinya = 0;

                    // Excelのwsに値を書き込む（geppoの行として）
                    wsGeppo.Cells[writeRow, normalMap.Day].Value = dbRow.B_Day;
                    wsGeppo.Cells[writeRow, normalMap.HansoCount].Value = hansoVal > 0 ? hansoVal : (object)null;
                    wsGeppo.Cells[writeRow, normalMap.YuryoKm].Value = yuryoKm > 0 ? yuryoKm : (object)null;
                    wsGeppo.Cells[writeRow, normalMap.MuryoKm].Value = muryoKm > 0 ? muryoKm : (object)null;

                    if (isOotsuki)
                        wsGeppo.Cells[writeRow, normalMap.ShinyaFee].Value = dbRow.H_LateFeeOotsuki > 0 ? dbRow.H_LateFeeOotsuki : (object)null;
                    else
                        wsGeppo.Cells[writeRow, normalMap.ShinyaMinutes].Value = dbRow.K_LateMinutes > 0 ? dbRow.K_LateMinutes : (object)null;

                    // フラグ書き込み
                    foreach (var flag in flags)
                    {
                        int? fv = dbRow.FlagValues?.GetValueOrDefault(flag.Id);
                        wsGeppo.Cells[writeRow, flag.ExcelColumn].Value = fv == 1 ? 1 : (object)null;
                    }

                    if (hansoVal > 0)
                    {
                        rowKihon = ratesForSheet.BaseFee;
                        if (yuryoKm > 0)
                            rowSoko = (Math.Floor(yuryoKm / 10) + 1) * ratesForSheet.MileageFee;

                        foreach (var flag in withAmountFlags)
                        {
                            bool flagOn = (dbRow.FlagValues?.GetValueOrDefault(flag.Id) == 1);
                            if (!flagOn) continue;

                            bool applyBase = flag.TargetFee == TargetFee.BaseFee || flag.TargetFee == TargetFee.Both;
                            bool applyMileage = flag.TargetFee == TargetFee.MileageFee || flag.TargetFee == TargetFee.Both;

                            if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                            {
                                if (applyBase) rowKihon = Math.Floor(ratesForSheet.BaseFee * flag.AmountValue.Value);
                                if (applyMileage) rowSoko = Math.Floor(rowSoko * flag.AmountValue.Value);
                            }
                            else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                            {
                                if (applyBase) rowKihon = flag.AmountValue.Value;
                                if (applyMileage) rowSoko = flag.AmountValue.Value;
                            }
                        }
                        if (isOotsuki)
                            rowShinya = dbRow.H_LateFeeOotsuki ?? 0;
                        else
                        {
                            double shinyaMin = dbRow.K_LateMinutes ?? 0;
                            if (shinyaMin > 0)
                            {
                                double numBlocks = Math.Floor(shinyaMin / 30) + 1;
                                double variableRyo = numBlocks * ratesForSheet.LateNightUnitFee;
                                rowShinya = variableRyo + ratesForSheet.LateNightFixedFee;
                            }
                        }
                    }

                    wsGeppo.Cells[writeRow, normalMap.KihonFee].Value = rowKihon > 0 ? rowKihon : (object)null;
                    wsGeppo.Cells[writeRow, normalMap.SokoFee].Value = rowSoko > 0 ? rowSoko : (object)null;
                    wsGeppo.Cells[writeRow, normalMap.ShinyaFee].Value = rowShinya > 0 ? rowShinya : (object)null;
                    double rowTotal = rowKihon + rowSoko + rowShinya;
                    wsGeppo.Cells[writeRow, normalMap.TotalFee].Value = rowTotal > 0 ? rowTotal : (object)null;

                    totalKihon += rowKihon;
                    totalSoko += rowSoko;
                    totalShinya += rowShinya;
                    totalSum += rowTotal;
                    writeRow++;
                }
            }
            else
            {
                // ── Excel使用時（従来）：wsInから直接読む ──
                for (int row = 3; row < totalRowIdx; row++)
                {
                    int hansoVal = GetInt(wsIn.Cells[row, normalMap.HansoCount].Value);
                    double rowKihon = 0, rowSoko = 0, rowShinya = 0;

                    if (hansoVal > 0)
                    {
                        double yuryoKmVal = GetDouble(wsIn.Cells[row, normalMap.YuryoKm].Value);

                        // 動的フラグによる基本料金計算
                        rowKihon = ratesForSheet.BaseFee;
                        if (yuryoKmVal > 0)
                            rowSoko = (Math.Floor(yuryoKmVal / 10) + 1) * ratesForSheet.MileageFee;

                        foreach (var flag in withAmountFlags)
                        {
                            bool flagOn = GetInt(wsIn.Cells[row, flag.ExcelColumn].Value) == 1;
                            if (!flagOn) continue;

                            bool applyBase = flag.TargetFee == TargetFee.BaseFee || flag.TargetFee == TargetFee.Both;
                            bool applyMileage = flag.TargetFee == TargetFee.MileageFee || flag.TargetFee == TargetFee.Both;

                            if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                            {
                                if (applyBase) rowKihon = Math.Floor(ratesForSheet.BaseFee * flag.AmountValue.Value);
                                if (applyMileage) rowSoko = Math.Floor(rowSoko * flag.AmountValue.Value);
                            }
                            else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                            {
                                if (applyBase) rowKihon = flag.AmountValue.Value;
                                if (applyMileage) rowSoko = flag.AmountValue.Value;
                            }
                        }

                        if (isOotsuki)
                        {
                            rowShinya = GetDouble(wsIn.Cells[row, normalMap.ShinyaFee].Value);
                        }
                        else
                        {
                            double shinyaMin = GetDouble(wsIn.Cells[row, normalMap.ShinyaMinutes].Value);
                            if (shinyaMin > 0)
                            {
                                double numBlocks = Math.Floor(shinyaMin / 30) + 1;
                                double variableRyo = numBlocks * ratesForSheet.LateNightUnitFee;
                                rowShinya = variableRyo + ratesForSheet.LateNightFixedFee;
                            }
                        }
                    }

                    wsGeppo.Cells[row, normalMap.KihonFee].Value = rowKihon > 0 ? rowKihon : null;
                    wsGeppo.Cells[row, normalMap.SokoFee].Value = rowSoko > 0 ? rowSoko : null;
                    wsGeppo.Cells[row, normalMap.ShinyaFee].Value = rowShinya > 0 ? rowShinya : null;
                    double rowTotal = rowKihon + rowSoko + rowShinya;
                    wsGeppo.Cells[row, normalMap.TotalFee].Value = rowTotal > 0 ? rowTotal : null;

                    totalKihon += rowKihon;
                    totalSoko += rowSoko;
                    totalShinya += rowShinya;
                    totalSum += rowTotal;
                }
            } // end Excel使用時

            wsGeppo.Cells[totalRowIdx, normalMap.KihonFee].Value = totalKihon > 0 ? totalKihon : null;
            wsGeppo.Cells[totalRowIdx, normalMap.SokoFee].Value = totalSoko > 0 ? totalSoko : null;
            wsGeppo.Cells[totalRowIdx, normalMap.ShinyaFee].Value = totalShinya > 0 ? totalShinya : null;
            wsGeppo.Cells[totalRowIdx, normalMap.TotalFee].Value = totalSum > 0 ? totalSum : null;

            var shukeiSheetName = wbShukei.Workbook.Worksheets.Any(ws => ws.Name == sheetName)
                ? sheetName
                : wbShukei.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.EndsWith(sheetName))?.Name;
            if (shukeiSheetName != null)
            {
                var wsShukei = wbShukei.Workbook.Worksheets[shukeiSheetName];
                var totals = CalculateTotals(wsIn, totalRowIdx, normalMap);
                wsShukei.Cells[shukeiMap.Days].Value = totals.days;
                wsShukei.Cells[shukeiMap.Hanso].Value = totals.hanso;
                wsShukei.Cells[shukeiMap.YuryoKm].Value = totals.yuryoKm;
                wsShukei.Cells[shukeiMap.MuryoKm].Value = totals.muryoKm;
                wsShukei.Cells[shukeiMap.Total].Value = totalSum > 0 ? totalSum : null;
            }
        }

        private (int days, int hanso, double yuryoKm, double muryoKm) CalculateTotals(ExcelWorksheet ws, int totalRowIdx, SheetColumnMap map)
        {
            int totalDays = 0, totalHanso = 0;
            double totalYuryoKm = 0, totalMuryoKm = 0;
            for (int row = 3; row < totalRowIdx; row++)
            {
                if (ws.Cells[row, map.Day].Value != null) totalDays++;
                totalHanso += GetInt(ws.Cells[row, map.HansoCount].Value);
                totalYuryoKm += GetDouble(ws.Cells[row, map.YuryoKm].Value);
                totalMuryoKm += GetDouble(ws.Cells[row, map.MuryoKm].Value);
            }
            return (totalDays, totalHanso, totalYuryoKm, totalMuryoKm);
        }

        private void ProcessEastSheet(ExcelPackage wbInput, ExcelPackage wbShukei, string sheetName, ColumnMapping columnMap)
        {
            if (wbShukei.Workbook.Worksheets.All(ws => ws.Name != sheetName)) return;
            var wsIn = wbInput.Workbook.Worksheets[sheetName];
            var wsShukei = wbShukei.Workbook.Worksheets[sheetName];
            var shukeiMap = columnMap.ShukeiSheet;
            wsShukei.Cells[shukeiMap.Days].Value = wsIn.Cells[shukeiMap.Days].Value;
            wsShukei.Cells[shukeiMap.Hanso].Value = wsIn.Cells[shukeiMap.Hanso].Value;
            wsShukei.Cells[shukeiMap.YuryoKm].Value = wsIn.Cells[shukeiMap.YuryoKm].Value;
            wsShukei.Cells[shukeiMap.MuryoKm].Value = wsIn.Cells[shukeiMap.MuryoKm].Value;
            wsShukei.Cells[shukeiMap.Total].Value = wsIn.Cells[shukeiMap.Total].Value;
            Logger.Info($"[{sheetName}] の値を転記しました。");
        }

        private int FindTotalRow(ExcelWorksheet ws)
        {
            if (ws?.Dimension == null) return -1;
            for (int row = ws.Dimension.End.Row; row >= 3; row--) { if (ws.Cells[row, 1].Value?.ToString()?.Contains("合計") == true) return row; }
            return -1;
        }

        private int GetInt(object val) => val == null ? 0 : (int)Convert.ToDouble(val);
        private double GetDouble(object val) => val == null ? 0.0 : Convert.ToDouble(val);
    }
}