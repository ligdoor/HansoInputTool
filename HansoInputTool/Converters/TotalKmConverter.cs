using System;
using System.Globalization;
using System.Windows.Data;
using HansoInputTool.Models;

namespace HansoInputTool.Converters
{
    /// <summary>
    /// Statisticsオブジェクトから総走行距離を計算するコンバーター
    /// </summary>
    public class TotalKmConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Statistics stats)
            {
                return stats.TotalYuryoKm + stats.TotalMuryoKm;
            }
            return 0.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}