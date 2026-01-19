using System;
using System.IO;
using System.Text.RegularExpressions;

namespace HansoInputTool.Models
{
    /// <summary>
    /// 月報ファイルの情報を保持するモデル
    /// </summary>
    public class MonthlyReportFile
    {
        /// <summary>
        /// ファイルの完全パス
        /// </summary>
        public string FilePath { get; set; }
        
        /// <summary>
        /// 親フォルダ名（サブフォルダ名）
        /// </summary>
        public string FolderName { get; set; }
        
        /// <summary>
        /// ファイル名
        /// </summary>
        public string FileName { get; set; }
        
        /// <summary>
        /// 期（例: 06）
        /// </summary>
        public int Period { get; set; }
        
        /// <summary>
        /// 月（例: 10）
        /// </summary>
        public int Month { get; set; }
        
        /// <summary>
        /// 年度（例: R6 → 2024）
        /// </summary>
        public int Year { get; set; }
        
        /// <summary>
        /// 表示用の月名（例: "2024年10月"）
        /// </summary>
        public string DisplayName => $"{Year}年{Month}月";
        
        /// <summary>
        /// ソート用のキー
        /// </summary>
        public string SortKey => $"{Year:D4}{Month:D2}";
        
        /// <summary>
        /// ファイル名から情報を解析
        /// </summary>
        public static MonthlyReportFile Parse(string filePath)
        {
            var fileName = Path.GetFileName(filePath);
            var folderName = Path.GetFileName(Path.GetDirectoryName(filePath));
            
            // ファイル名パターン: ##期#月R#アルス搬送・霊柩車　実績月報.xls
            var pattern = @"(\d+)期(\d+)月R(\d+)アルス搬送・霊柩車\s*実績月報\.xls";
            var match = Regex.Match(fileName, pattern);
            
            if (!match.Success)
            {
                return null;
            }
            
            var period = int.Parse(match.Groups[1].Value);
            var month = int.Parse(match.Groups[2].Value);
            var reiwaYear = int.Parse(match.Groups[3].Value);
            
            // 令和年を西暦に変換（令和元年 = 2019年）
            var year = 2018 + reiwaYear;
            
            return new MonthlyReportFile
            {
                FilePath = filePath,
                FolderName = folderName,
                FileName = fileName,
                Period = period,
                Month = month,
                Year = year
            };
        }
        
        /// <summary>
        /// ファイルが存在するか確認
        /// </summary>
        public bool Exists()
        {
            return File.Exists(FilePath);
        }
        
        public override string ToString()
        {
            return DisplayName;
        }
    }
}
