using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 月報ファイルをスキャンして検出するサービス
    /// </summary>
    public class MonthlyReportScanner
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        /// <summary>
        /// 指定されたフォルダ配下から月報ファイルを検索
        /// </summary>
        /// <param name="rootFolder">ルートフォルダパス</param>
        /// <returns>検出された月報ファイルのリスト</returns>
        public List<MonthlyReportFile> ScanMonthlyReports(string rootFolder)
        {
            var reports = new List<MonthlyReportFile>();
            
            try
            {
                Logger.Info($"月報ファイルをスキャン中: {rootFolder}");
                
                if (!Directory.Exists(rootFolder))
                {
                    Logger.Warn($"指定されたフォルダが存在しません: {rootFolder}");
                    return reports;
                }
                
                // サブフォルダを取得
                var subFolders = Directory.GetDirectories(rootFolder);
                Logger.Info($"サブフォルダ数: {subFolders.Length}");
                
                foreach (var subFolder in subFolders)
                {
                    try
                    {
                        // サブフォルダ内のxlsファイルを検索
                        var xlsFiles = Directory.GetFiles(subFolder, "*.xls", SearchOption.TopDirectoryOnly);
                        
                        foreach (var xlsFile in xlsFiles)
                        {
                            var fileName = Path.GetFileName(xlsFile);
                            
                            // パターンに一致するファイルのみ処理
                            if (IsMonthlyReportFile(fileName))
                            {
                                var report = MonthlyReportFile.Parse(xlsFile);
                                
                                if (report != null)
                                {
                                    reports.Add(report);
                                    Logger.Info($"月報ファイル検出: {report.DisplayName} - {fileName}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(ex, $"サブフォルダのスキャン中にエラー: {subFolder}");
                    }
                }
                
                // 年月でソート
                reports = reports.OrderBy(r => r.SortKey).ToList();
                
                Logger.Info($"月報ファイルスキャン完了: {reports.Count}件");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "月報ファイルのスキャン中にエラーが発生");
            }
            
            return reports;
        }
        
        /// <summary>
        /// ファイル名が月報ファイルのパターンに一致するか確認
        /// </summary>
        private bool IsMonthlyReportFile(string fileName)
        {
            // パターン: ##期#月R#アルス搬送・霊柩車　実績月報.xls
            var pattern = @"^\d+期\d+月R\d+アルス搬送・霊柩車\s*実績月報\.xls$";
            return Regex.IsMatch(fileName, pattern);
        }
        
        /// <summary>
        /// 年度でグループ化
        /// </summary>
        public Dictionary<int, List<MonthlyReportFile>> GroupByYear(List<MonthlyReportFile> reports)
        {
            return reports
                .GroupBy(r => r.Year)
                .ToDictionary(g => g.Key, g => g.OrderBy(r => r.Month).ToList());
        }
        
        /// <summary>
        /// 期でグループ化
        /// </summary>
        public Dictionary<int, List<MonthlyReportFile>> GroupByPeriod(List<MonthlyReportFile> reports)
        {
            return reports
                .GroupBy(r => r.Period)
                .ToDictionary(g => g.Key, g => g.OrderBy(r => r.Year).ThenBy(r => r.Month).ToList());
        }
        
        /// <summary>
        /// 最新の年度を取得
        /// </summary>
        public int GetLatestYear(List<MonthlyReportFile> reports)
        {
            return reports.Any() ? reports.Max(r => r.Year) : DateTime.Now.Year;
        }
        
        /// <summary>
        /// 指定年度の月報ファイルを取得
        /// </summary>
        public List<MonthlyReportFile> GetReportsByYear(List<MonthlyReportFile> reports, int year)
        {
            return reports
                .Where(r => r.Year == year)
                .OrderBy(r => r.Month)
                .ToList();
        }
    }
}
