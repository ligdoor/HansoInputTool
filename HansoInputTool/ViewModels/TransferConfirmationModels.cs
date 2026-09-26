// ViewModels/TransferConfirmationViewModel.cs
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace HansoInputTool.ViewModels
{
    public class VehicleTab : ObservableObject
    {
        public string SheetName { get; set; }
        public string DisplayName { get; set; }
        public int WorkingDays { get; set; }
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public double EstimatedRevenue { get; set; }
        public SolidColorBrush StatusColor { get; set; }
        public string StatusIcon { get; set; }

        public string TabHeader => $"{StatusIcon} {DisplayName}";
        public string WorkingDaysText => $"{WorkingDays}日稼働";
        public string IssueCountText
        {
            get
            {
                if (ErrorCount > 0)
                    return $"エラー: {ErrorCount}件";
                if (WarningCount > 0)
                    return $"警告: {WarningCount}件";
                return "問題なし";
            }
        }
    }

    public class TransferRowData
    {
        public int RowIndex { get; set; }
        public string Day { get; set; }
        public string HansoCount { get; set; }
        public string YuryoKm { get; set; }
        public string MuryoKm { get; set; }
        public string IsKoryo { get; set; }
        public string LateValue { get; set; }
        public string BaseFee { get; set; }
        public string MileageFee { get; set; }
        public string LateFee { get; set; }
        public string TotalFee { get; set; }

        public bool HasError { get; set; }
        public bool HasWarning { get; set; }
    }

    public class VehicleSummary
    {
        public int WorkingDays { get; set; }
        public int TotalHanso { get; set; }
        public double TotalYuryoKm { get; set; }
        public double TotalMuryoKm { get; set; }
        public double TotalKm { get; set; }
        public int KoryoCount { get; set; }
        public double EstimatedRevenue { get; set; }
        public double AverageRevenuePerTrip { get; set; }
    }

    public class ValidationIssue
    {
        public IssueSeverity Severity { get; set; }
        public string SheetName { get; set; }
        public int Day { get; set; }
        public string Message { get; set; }
        public string Icon { get; set; }

        public SolidColorBrush Color
        {
            get
            {
                return Severity switch
                {
                    IssueSeverity.Error => new SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68)),
                    IssueSeverity.Warning => new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 191, 36)),
                    IssueSeverity.Info => new SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 130, 246)),
                    _ => Brushes.Gray
                };
            }
        }
    }

    public enum IssueSeverity
    {
        Info,
        Warning,
        Error
    }
}
