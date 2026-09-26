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
            var lateMinutes = row.K_LateMinutes ?? 0;

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

            if (!isOotsuki && lateMinutes > 180)
            {
                issues.Add(new ValidationIssue
                {
                    Severity = IssueSeverity.Warning,
                    SheetName = sheetName,
                    Day = day,
                    Message = $"{day}日目: 深夜時間が3時間を超えています({lateMinutes}分)",
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
