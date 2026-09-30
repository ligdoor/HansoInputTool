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
    /// <summary>
    /// TransferConfirmationViewModel の分割定義：行データのバリデーション（エラー・警告判定）まわり。
    /// </summary>
    public partial class TransferConfirmationViewModel
    {
        private List<ValidationIssue> ValidateRow(RowData row, string sheetName, bool isOotsuki)
        {
            var issues = new List<ValidationIssue>();

            if (!row.B_Day.HasValue) return issues;

            var day = row.B_Day.Value;
            var yuryoKm = row.D_YuryoKm ?? 0;
            var muryoKm = row.E_MuryoKm ?? 0;
            bool hasLateValueInWrongMode = isOotsuki
                ? row.K_LateMinutes.HasValue
                : row.H_LateFeeOotsuki.HasValue;

            if (hasLateValueInWrongMode)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Error,
                    SheetName = sheetName,
                    Day = day,
                    Message = isOotsuki
                        ? $"{day}日目: 車両設定は深夜料金ですが、深夜時間欄に値があります"
                        : $"{day}日目: 車両設定は深夜時間ですが、深夜料金欄に値があります",
                    Icon = "❌"
                });
            }

            double lateValue = isOotsuki
                ? row.H_LateFeeOotsuki.GetValueOrDefault()
                : row.K_LateMinutes.GetValueOrDefault();
            string lateFieldName = isOotsuki ? "深夜料金" : "深夜時間";

            // エラーチェック
            if (yuryoKm < 1 && row.C_Hanso > 0)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Error,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 搬送があるのに有料キロが0または未入力です",
                    Icon = "❌"
                });
            }

            if (yuryoKm > 500)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Error,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 有料キロが異常に多い({yuryoKm}km) - 入力ミスの可能性",
                    Icon = "❌"
                });
            }

            // 警告チェック
            if (yuryoKm > 300)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Warning,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 有料キロが通常より長い({yuryoKm}km)",
                    Icon = "⚠️"
                });
            }

            if (lateValue < 0)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Error,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: {lateFieldName}が0未満です({lateValue})",
                    Icon = "❌"
                });
            }
            else if (isOotsuki && lateValue > 50000)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Warning,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 深夜料金が50,000円を超えています({lateValue:N0}円)",
                    Icon = "⚠️"
                });
            }
            else if (!isOotsuki && lateValue > 1440)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Error,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 深夜時間が24時間を超えています({lateValue}分)",
                    Icon = "❌"
                });
            }
            else if (!isOotsuki && lateValue > 720)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Warning,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 深夜時間が12時間を超えています({lateValue}分)",
                    Icon = "⚠️"
                });
            }
            else if (!isOotsuki && lateValue > 180)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Warning,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 深夜時間が3時間を超えています({lateValue}分)",
                    Icon = "⚠️"
                });
            }

            if (muryoKm > yuryoKm && yuryoKm > 0)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Info,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 無料キロが有料キロを上回っています",
                    Icon = "ℹ️"
                });
            }

            return issues;
        }

        private void FilterValidationIssues()
        {
            ValidationIssues.Clear();

            var filtered = ShowErrorsOnly
                ? _allValidationIssues.Where(i => i.Severity == IssueSeverity.Error)
                : _allValidationIssues;

            foreach (var issue in filtered)
            {
                ValidationIssues.Add(issue);
            }
        }
    }
}
