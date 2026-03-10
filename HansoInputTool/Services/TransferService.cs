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
            IProgress<TransferProgressReport> progress)
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

                progress.Report(new TransferProgressReport { Message = "--- 全シートの転記処理を開始 ---" });
                Logger.Info("--- 全シートの転記処理を開始 ---");

                var sheetsToProcess = allSheetNames?.Where(s => !s.Contains("登録")).ToList() ?? new List<string>();
                int totalSheets = sheetsToProcess.Count;
                int processedCount = 0;

                foreach (var sheetName in sheetsToProcess)
                {
                    progress.Report(new TransferProgressReport { Current = processedCount, Total = totalSheets, Message = $"処理中: {sheetName} ..." });

                    if (IsNormalSheet(sheetName))
                    {
                        ProcessNormalSheet(wbInput, wbGeppo, wbShukei, sheetName, rates, columnMap);
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
                    Logger.Info($"[{sheetName}] の処理が完了しました。");
                }

                progress.Report(new TransferProgressReport { Current = processedCount, Total = totalSheets, Message = "最終処理中..." });

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
            });
        }

        // ===== シート種別の一元判定ヘルパー =====
        private static readonly string[] NormalSheetKeywords = { "寝台車", "霊柩車", "CH富士吉田", "CH大月", "CH東富士" };
        private static readonly string[] EastSheetKeywords  = { "東日本セレモニー", "東日本" };

        private static bool IsNormalSheet(string sheetName) =>
            NormalSheetKeywords.Any(kw => sheetName.Contains(kw));

        private static bool IsEastSheet(string sheetName) =>
            EastSheetKeywords.Any(kw => sheetName.Contains(kw));
        // ==========================================

        private void ProcessNormalSheet(ExcelPackage wbInput, ExcelPackage wbGeppo, ExcelPackage wbShukei, string sheetName, Dictionary<string, RateInfo> rates, ColumnMapping columnMap)
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

            for (int row = 3; row < totalRowIdx; row++)
            {
                int hansoVal = GetInt(wsIn.Cells[row, normalMap.HansoCount].Value);
                double rowKihon = 0, rowSoko = 0, rowShinya = 0;

                if (hansoVal > 0)
                {
                    double yuryoKmVal = GetDouble(wsIn.Cells[row, normalMap.YuryoKm].Value);
                    bool isKoryo = GetInt(wsIn.Cells[row, normalMap.IsKoryo].Value) == 1;

                    rowKihon = isKoryo ? Math.Floor((double)ratesForSheet.BaseFee / 2) : ratesForSheet.BaseFee;

                    if (yuryoKmVal > 0)
                    {
                        rowSoko = (Math.Floor(yuryoKmVal / 10) + 1) * ratesForSheet.MileageFee;
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

            wsGeppo.Cells[totalRowIdx, normalMap.KihonFee].Value = totalKihon > 0 ? totalKihon : null;
            wsGeppo.Cells[totalRowIdx, normalMap.SokoFee].Value = totalSoko > 0 ? totalSoko : null;
            wsGeppo.Cells[totalRowIdx, normalMap.ShinyaFee].Value = totalShinya > 0 ? totalShinya : null;
            wsGeppo.Cells[totalRowIdx, normalMap.TotalFee].Value = totalSum > 0 ? totalSum : null;

            if (wbShukei.Workbook.Worksheets.Any(ws => ws.Name == sheetName))
            {
                var wsShukei = wbShukei.Workbook.Worksheets[sheetName];
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