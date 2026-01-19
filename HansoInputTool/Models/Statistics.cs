using System;

namespace HansoInputTool.Models
{
    /// <summary>
    /// 統計情報を格納するモデル
    /// </summary>
    public class Statistics
    {
        // 月次統計
        public int TotalHanso { get; set; }                    // 総搬送回数
        public double TotalYuryoKm { get; set; }               // 総有料キロ
        public double TotalMuryoKm { get; set; }               // 総無料キロ
        public double AverageYuryoKm { get; set; }             // 平均有料キロ
        public double AverageMuryoKm { get; set; }             // 平均無料キロ
        public int TotalKoryo { get; set; }                    // 行旅回数
        public double TotalLateCharges { get; set; }           // 総深夜料金
        
        // 売上関連
        public double EstimatedRevenue { get; set; }           // 推定売上
        public double AverageRevenuePerTrip { get; set; }      // 1回あたり平均売上
        
        // 車両別統計
        public int ActiveVehicleCount { get; set; }            // 稼働車両数
        public string MostUsedVehicle { get; set; }            // 最多使用車両
        public int MostUsedVehicleCount { get; set; }          // 最多使用車両の使用回数
        
        // 日別統計
        public int MaxDailyHanso { get; set; }                 // 1日最大搬送回数
        public DateTime MaxDailyHansoDate { get; set; }        // 最大搬送回数の日付
        public double MaxDailyKm { get; set; }                 // 1日最大走行距離
        public DateTime MaxDailyKmDate { get; set; }           // 最大走行距離の日付
        
        // 営業日数
        public int WorkingDays { get; set; }                   // 営業日数
        public double AverageHansoPerDay { get; set; }         // 1日あたり平均搬送回数
        
        // トレンド情報（前月比）
        public double HansoChangePercent { get; set; }         // 搬送回数変化率
        public double RevenueChangePercent { get; set; }       // 売上変化率
        public double KmChangePercent { get; set; }            // 走行距離変化率
        
        public Statistics()
        {
            MostUsedVehicle = "N/A";
            MaxDailyHansoDate = DateTime.Now;
            MaxDailyKmDate = DateTime.Now;
        }
    }
}
