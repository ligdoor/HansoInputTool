using System;
using HansoInputTool.Models;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 会社出発時刻〜故人宅出発時刻の区間から、深夜割増時間帯（22:00〜翌5:00）に該当する分数と、
    /// それに対応する深夜料金を計算する。
    /// 深夜料金の計算式（30分単位切り上げ）は Services/TransferService.cs・
    /// ViewModels/TransferConfirmationViewModel.cs の本処理と同じ式を使用している。
    /// この式を変更する場合は、必ず両方に同じ変更を反映すること。
    /// </summary>
    public static class LateNightFeeCalculator
    {
        private static readonly TimeSpan NightStart = new TimeSpan(22, 0, 0); // 22:00
        private static readonly TimeSpan NightSpan = TimeSpan.FromHours(7);   // 22:00〜翌5:00 = 7時間

        /// <summary>
        /// 出発時刻(start)〜出発時刻(end)の区間のうち、22:00〜翌5:00に該当する分数を返す。
        /// endがstartと同時刻または時刻的に早い場合は、日をまたいだ（翌日扱い）区間として計算する。
        /// </summary>
        public static int CalculateNightMinutes(TimeSpan start, TimeSpan end)
        {
            var baseDate = DateTime.Today;
            var startDt = baseDate + start;
            var endDt = baseDate + end;
            if (endDt <= startDt) endDt = endDt.AddDays(1);

            int nightMinutes = 0;

            // 区間をカバーしうる深夜帯（前日22:00〜翌5:00）を日ごとに列挙し、区間との重なりを合算する。
            // 実運用では区間が数日にまたがることは想定していないが、念のため区間全体をカバーするまで繰り返す。
            var windowStart = startDt.Date.AddDays(-1) + NightStart;
            while (windowStart < endDt)
            {
                var windowEnd = windowStart + NightSpan;
                var overlapStart = startDt > windowStart ? startDt : windowStart;
                var overlapEnd = endDt < windowEnd ? endDt : windowEnd;
                if (overlapEnd > overlapStart)
                    nightMinutes += (int)Math.Round((overlapEnd - overlapStart).TotalMinutes);
                windowStart = windowStart.AddDays(1);
            }

            return nightMinutes;
        }

        /// <summary>30分単位切り上げでのブロック数（1〜30分=1ブロック、31〜60分=2ブロック…）。0分なら0。</summary>
        public static int CalculateBlocks(int minutes)
            => minutes <= 0 ? 0 : (int)(Math.Floor((double)minutes / 30) + 1);

        /// <summary>深夜料金（円）＝深夜固定＋ブロック数×深夜単価。rateが無い、または0分なら0。</summary>
        public static int CalculateFee(int minutes, RateInfo rate)
        {
            if (rate == null || minutes <= 0) return 0;
            int blocks = CalculateBlocks(minutes);
            return rate.LateNightFixedFee + (blocks * rate.LateNightUnitFee);
        }
    }
}
