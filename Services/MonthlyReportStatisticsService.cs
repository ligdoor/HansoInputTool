using System;
using System.Collections.Generic;
using System.Linq;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;
using System.IO;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 月報ファイルから統計情報を計算するサービス
    /// </summary>
    public class MonthlyReportStatisticsService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // 料金計算用の定数
        private const double YuryoKmRate = 150.0;
        private const double KoryoFee = 5000.0;

        // 内部データクラス（行旅情報を含む）
        private class InternalRowData
        {
            public int Day { get; set; }
            public double YuryoKm { get; set; }
            public double MuryoKm { get; set; }
            public double LateCharge { get; set; }
            public bool IsKoryo { get; set; }
        }

        /// <summary>
        /// 単一の月報ファイルから統計を計算
        /// </summary>
        public Statistics CalculateStatisticsFromFile(MonthlyReportFile reportFile)
        {
            try
            {
                Logger.Info($"統計計算開始: {reportFile.DisplayName}");

                if (!reportFile.Exists())
                {
                    Logger.Warn($"ファイルが存在しません: {reportFile.FilePath}");
                    return new Statistics();
                }

                // EPPlus 8の新しいライセンス設定方法
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(reportFile.FilePath)))
                {
                    var stats = new Statistics();
                    var allData = new List<InternalRowData>();
                    var vehicleUsageCount = new Dictionary<string, int>();

                    var vehicleSheets = GetVehicleSheets(package);
                    Logger.Info($"車両シート数: {vehicleSheets.Count}");

                    foreach (var sheetName in vehicleSheets)
                    {
                        try
                        {
                            var worksheet = package.Workbook.Worksheets[sheetName];
                            var sheetData = ReadSheetData(worksheet);

                            if (sheetData.Any())
                            {
                                allData.AddRange(sheetData);
                                vehicleUsageCount[sheetName] = sheetData.Count;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn(ex, $"シート {sheetName} の読み込みエラー");
                        }
                    }

                    if (!allData.Any())
                    {
                        Logger.Warn("データが見つかりませんでした");
                        return stats;
                    }

                    CalculateBasicStatistics(stats, allData, vehicleUsageCount);

                    Logger.Info($"統計計算完了: 総搬送回数={stats.TotalHanso}");
                    return stats;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"統計計算エラー: {reportFile.FilePath}");
                return new Statistics();
            }
        }

        /// <summary>
        /// 複数の月報ファイルから合算統計を計算
        /// </summary>
        public Statistics CalculateCombinedStatistics(List<MonthlyReportFile> reportFiles)
        {
            try
            {
                Logger.Info($"合算統計計算開始: {reportFiles.Count}ファイル");

                var combinedStats = new Statistics();
                var allMonthlyStats = new List<Statistics>();

                foreach (var file in reportFiles)
                {
                    var monthStats = CalculateStatisticsFromFile(file);
                    if (monthStats.TotalHanso > 0)
                    {
                        allMonthlyStats.Add(monthStats);
                    }
                }

                if (!allMonthlyStats.Any())
                {
                    Logger.Warn("有効な統計データが見つかりませんでした");
                    return combinedStats;
                }

                // 合算処理
                combinedStats.TotalHanso = allMonthlyStats.Sum(s => s.TotalHanso);
                combinedStats.TotalYuryoKm = allMonthlyStats.Sum(s => s.TotalYuryoKm);
                combinedStats.TotalMuryoKm = allMonthlyStats.Sum(s => s.TotalMuryoKm);
                combinedStats.TotalKoryo = allMonthlyStats.Sum(s => s.TotalKoryo);
                combinedStats.TotalLateCharges = allMonthlyStats.Sum(s => s.TotalLateCharges);
                combinedStats.EstimatedRevenue = allMonthlyStats.Sum(s => s.EstimatedRevenue);
                combinedStats.WorkingDays = allMonthlyStats.Sum(s => s.WorkingDays);

                // 平均値の計算
                combinedStats.AverageYuryoKm = combinedStats.TotalHanso > 0
                    ? combinedStats.TotalYuryoKm / combinedStats.TotalHanso : 0;
                combinedStats.AverageMuryoKm = combinedStats.TotalHanso > 0
                    ? combinedStats.TotalMuryoKm / combinedStats.TotalHanso : 0;
                combinedStats.AverageRevenuePerTrip = combinedStats.TotalHanso > 0
                    ? combinedStats.EstimatedRevenue / combinedStats.TotalHanso : 0;
                combinedStats.AverageHansoPerDay = combinedStats.WorkingDays > 0
                    ? (double)combinedStats.TotalHanso / combinedStats.WorkingDays : 0;

                // 車両統計（重複を除いて合計）
                var allVehicles = new HashSet<string>();
                foreach (var stat in allMonthlyStats)
                {
                    if (!string.IsNullOrEmpty(stat.MostUsedVehicle))
                    {
                        allVehicles.Add(stat.MostUsedVehicle);
                    }
                }
                combinedStats.ActiveVehicleCount = allVehicles.Count;

                // 最大値
                combinedStats.MaxDailyHanso = allMonthlyStats.Max(s => s.MaxDailyHanso);
                combinedStats.MaxDailyKm = allMonthlyStats.Max(s => s.MaxDailyKm);

                // 最大値の日付
                var maxHansoMonth = allMonthlyStats.OrderByDescending(s => s.MaxDailyHanso).FirstOrDefault();
                if (maxHansoMonth != null)
                {
                    combinedStats.MaxDailyHansoDate = maxHansoMonth.MaxDailyHansoDate;
                }

                var maxKmMonth = allMonthlyStats.OrderByDescending(s => s.MaxDailyKm).FirstOrDefault();
                if (maxKmMonth != null)
                {
                    combinedStats.MaxDailyKmDate = maxKmMonth.MaxDailyKmDate;
                }

                // 最多使用車両（全月で最も使用回数が多い車両）
                var vehicleDict = new Dictionary<string, int>();
                foreach (var stat in allMonthlyStats)
                {
                    if (!string.IsNullOrEmpty(stat.MostUsedVehicle))
                    {
                        if (!vehicleDict.ContainsKey(stat.MostUsedVehicle))
                            vehicleDict[stat.MostUsedVehicle] = 0;
                        vehicleDict[stat.MostUsedVehicle] += stat.MostUsedVehicleCount;
                    }
                }

                if (vehicleDict.Any())
                {
                    var mostUsed = vehicleDict.OrderByDescending(kvp => kvp.Value).First();
                    combinedStats.MostUsedVehicle = mostUsed.Key;
                    combinedStats.MostUsedVehicleCount = mostUsed.Value;
                }

                Logger.Info($"合算統計計算完了: 総搬送回数={combinedStats.TotalHanso}");
                return combinedStats;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "合算統計計算エラー");
                return new Statistics();
            }
        }

        /// <summary>
        /// 車両シート名を取得
        /// </summary>
        private List<string> GetVehicleSheets(ExcelPackage package)
        {
            var vehicleSheets = new List<string>();

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var sheetName = worksheet.Name;

                // 除外するシート名
                if (sheetName == "月間集計" ||
                    sheetName == "Template" ||
                    sheetName.StartsWith("_"))
                {
                    continue;
                }

                vehicleSheets.Add(sheetName);
            }

            return vehicleSheets;
        }

        /// <summary>
        /// シートからデータを読み込み
        /// </summary>
        private List<InternalRowData> ReadSheetData(ExcelWorksheet worksheet)
        {
            var data = new List<InternalRowData>();

            try
            {
                for (int row = 3; row <= worksheet.Dimension?.End.Row; row++)
                {
                    var dayValue = worksheet.Cells[row, 2].Value;
                    if (dayValue == null || string.IsNullOrWhiteSpace(dayValue.ToString()))
                        break;

                    var day = ParseInt(dayValue);
                    if (!day.HasValue)
                        continue;

                    var yuryoKm = ParseDouble(worksheet.Cells[row, 4].Value) ?? 0;
                    var muryoKm = ParseDouble(worksheet.Cells[row, 5].Value) ?? 0;
                    var lateText = worksheet.Cells[row, 6].Value?.ToString() ?? "";
                    var koryoText = worksheet.Cells[row, 7].Value?.ToString() ?? "";

                    var rowData = new InternalRowData
                    {
                        Day = day.Value,
                        YuryoKm = yuryoKm,
                        MuryoKm = muryoKm,
                        LateCharge = ParseLateValue(lateText),
                        IsKoryo = koryoText == "行旅"
                    };

                    data.Add(rowData);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, $"シートデータ読み込みエラー: {worksheet.Name}");
            }

            return data;
        }

        /// <summary>
        /// 基本統計を計算
        /// </summary>
        private void CalculateBasicStatistics(Statistics stats, List<InternalRowData> allData, Dictionary<string, int> vehicleUsageCount)
        {
            stats.TotalHanso = allData.Count;
            stats.TotalYuryoKm = allData.Sum(d => d.YuryoKm);
            stats.TotalMuryoKm = allData.Sum(d => d.MuryoKm);
            stats.AverageYuryoKm = stats.TotalHanso > 0 ? stats.TotalYuryoKm / stats.TotalHanso : 0;
            stats.AverageMuryoKm = stats.TotalHanso > 0 ? stats.TotalMuryoKm / stats.TotalHanso : 0;

            stats.TotalKoryo = allData.Count(d => d.IsKoryo);
            stats.TotalLateCharges = allData.Sum(d => d.LateCharge);

            stats.EstimatedRevenue =
                (stats.TotalYuryoKm * YuryoKmRate) +
                (stats.TotalKoryo * KoryoFee) +
                stats.TotalLateCharges;

            stats.AverageRevenuePerTrip = stats.TotalHanso > 0
                ? stats.EstimatedRevenue / stats.TotalHanso
                : 0;

            stats.ActiveVehicleCount = vehicleUsageCount.Count;
            if (vehicleUsageCount.Any())
            {
                var mostUsed = vehicleUsageCount.OrderByDescending(kvp => kvp.Value).First();
                stats.MostUsedVehicle = mostUsed.Key;
                stats.MostUsedVehicleCount = mostUsed.Value;
            }

            var dailyStats = allData
                .GroupBy(d => d.Day)
                .Select(g => new
                {
                    Day = g.Key,
                    Count = g.Count(),
                    TotalKm = g.Sum(d => d.YuryoKm + d.MuryoKm)
                })
                .ToList();

            if (dailyStats.Any())
            {
                var maxHansoDay = dailyStats.OrderByDescending(d => d.Count).First();
                stats.MaxDailyHanso = maxHansoDay.Count;
                stats.MaxDailyHansoDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, maxHansoDay.Day);

                var maxKmDay = dailyStats.OrderByDescending(d => d.TotalKm).First();
                stats.MaxDailyKm = maxKmDay.TotalKm;
                stats.MaxDailyKmDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, maxKmDay.Day);

                stats.WorkingDays = dailyStats.Count;
                stats.AverageHansoPerDay = (double)stats.TotalHanso / stats.WorkingDays;
            }
        }

        private int? ParseInt(object value)
        {
            if (value == null) return null;
            if (int.TryParse(value.ToString(), out int result))
                return result;
            if (double.TryParse(value.ToString(), out double dResult))
                return (int)dResult;
            return null;
        }

        private double? ParseDouble(object value)
        {
            if (value == null) return null;
            if (double.TryParse(value.ToString(), out double result))
                return result;
            return null;
        }

        private double ParseLateValue(string lateValueText)
        {
            if (string.IsNullOrWhiteSpace(lateValueText))
                return 0;

            var cleaned = lateValueText
                .Replace("￥", "")
                .Replace("¥", "")
                .Replace(",", "")
                .Trim();

            if (double.TryParse(cleaned, out double value))
                return value;

            return 0;
        }
    }
}