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
    /// TransferConfirmationViewModel の分割定義：車両データの読み込み・解析まわり。
    /// </summary>
    public partial class TransferConfirmationViewModel
    {
        private void LoadAllVehicles()
        {
            Logger.Info("転記確認: 全車両データを読み込み中...");

            var vehicleSheets = _excelHandler.GetVehicleSheetNames();
            int totalErrors = 0;
            int totalWarnings = 0;
            double totalRevenue = 0;

            foreach (var sheetName in vehicleSheets)
            {
                var vehicleData = AnalyzeVehicleData(sheetName);
                VehicleTabs.Add(vehicleData);

                totalErrors += vehicleData.ErrorCount;
                totalWarnings += vehicleData.WarningCount;
                totalRevenue += vehicleData.EstimatedRevenue;
            }

            // 統計情報
            TotalVehicles = vehicleSheets.Count.ToString();
            VehiclesWithErrors = VehicleTabs.Count(v => v.ErrorCount > 0).ToString();
            VehiclesWithWarnings = VehicleTabs.Count(v => v.WarningCount > 0).ToString();
            TotalEstimatedRevenue = totalRevenue.ToString("N0");

            OnPropertyChanged(nameof(TotalVehicles));
            OnPropertyChanged(nameof(VehiclesWithErrors));
            OnPropertyChanged(nameof(VehiclesWithWarnings));
            OnPropertyChanged(nameof(TotalEstimatedRevenue));

            // 最初の車両を選択（エラーがある車両を優先）
            SelectedVehicle = VehicleTabs.FirstOrDefault(v => v.ErrorCount > 0)
                           ?? VehicleTabs.FirstOrDefault();

            Logger.Info($"転記確認: {vehicleSheets.Count}台分のデータを読み込み完了");
            Logger.Info($"エラー: {totalErrors}件、警告: {totalWarnings}件");
        }

        private VehicleTab AnalyzeVehicleData(string sheetName)
        {
            var rows = _excelHandler.GetSheetDataForPreview(sheetName);
            // [深夜料金バグ修正] 車両設定（深夜入力方式）を優先する
            bool isOotsuki = _excelHandler.IsFeeMode(sheetName);
            string rateCategory = sheetName.Contains("霊柩車") ? "霊柩車" : "寝台車";

            if (!_rates.TryGetValue(rateCategory, out var rate))
            {
                Logger.Warn($"料金カテゴリが見つかりません: {rateCategory}");
                rate = _rates.Values.FirstOrDefault();
            }

            var vehicleTab = new VehicleTab
            {
                SheetName = sheetName,
                DisplayName = sheetName,
                WorkingDays = rows.Count(r => r.B_Day.HasValue)
            };

            int errorCount = 0;
            int warningCount = 0;
            double totalRevenue = 0;

            foreach (var row in rows)
            {
                if (!row.B_Day.HasValue) continue;

                // 金額ありフラグを動的に取得
                var withAmountFlags = _flagService?.Flags
                    .Where(f => f.Type == FlagType.WithAmount).ToList()
                    ?? new List<FlagDefinition>();

                // 料金計算
                double baseFee = rate.BaseFee;
                foreach (var flag in withAmountFlags)
                {
                    if (!row.GetFlag(flag.Id)) continue;
                    if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                        baseFee = Math.Floor(rate.BaseFee * flag.AmountValue.Value);
                    else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                        baseFee = flag.AmountValue.Value;
                }
                double mileageFee = 0;
                double lateFee = 0;

                if (row.D_YuryoKm.HasValue && row.D_YuryoKm > 0)
                {
                    mileageFee = (Math.Floor((double)row.D_YuryoKm / 10) + 1) * rate.MileageFee;
                }

                if (isOotsuki && row.H_LateFeeOotsuki.HasValue)
                {
                    lateFee = row.H_LateFeeOotsuki.Value;
                }
                else if (!isOotsuki && row.K_LateMinutes.HasValue && row.K_LateMinutes > 0)
                {
                    double blocks = Math.Floor((double)row.K_LateMinutes / 30) + 1;
                    lateFee = rate.LateNightFixedFee + (blocks * rate.LateNightUnitFee);
                }

                double totalFee = baseFee + mileageFee + lateFee;
                totalRevenue += totalFee;

                // バリデーション
                var issues = ValidateRow(row, sheetName, isOotsuki);
                errorCount += issues.Count(i => i.Severity == IssueSeverity.Error);
                warningCount += issues.Count(i => i.Severity == IssueSeverity.Warning);
            }

            vehicleTab.ErrorCount = errorCount;
            vehicleTab.WarningCount = warningCount;
            vehicleTab.EstimatedRevenue = totalRevenue;
            vehicleTab.StatusColor = GetStatusColor(errorCount, warningCount);
            vehicleTab.StatusIcon = GetStatusIcon(errorCount, warningCount);

            return vehicleTab;
        }

        private void LoadVehicleData()
        {
            if (SelectedVehicle == null) return;

            Logger.Info($"車両データ読み込み: {SelectedVehicle.SheetName}");

            CurrentVehicleRows.Clear();
            _allValidationIssues.Clear();

            var rows = _excelHandler.GetSheetDataForPreview(SelectedVehicle.SheetName);
            // [深夜料金バグ修正] 車両設定（深夜入力方式）を優先する
            bool isOotsuki = _excelHandler.IsFeeMode(SelectedVehicle.SheetName);
            string rateCategory = SelectedVehicle.SheetName.Contains("霊柩車") ? "霊柩車" : "寝台車";

            if (!_rates.TryGetValue(rateCategory, out var rate))
            {
                rate = _rates.Values.FirstOrDefault();
            }

            double totalYuryoKm = 0;
            double totalMuryoKm = 0;
            double totalRevenue = 0;
            int totalHanso = 0;
            int koryoCount = 0;

            foreach (var row in rows)
            {
                // 料金計算
                double baseFee = 0;
                double mileageFee = 0;
                double lateFee = 0;
                double totalFee = 0;

                if (row.B_Day.HasValue)
                {
                    totalHanso   += row.C_Hanso    ?? 0;
                    totalYuryoKm += row.D_YuryoKm  ?? 0;
                    totalMuryoKm += row.E_MuryoKm  ?? 0;

                    // 金額ありフラグを動的に取得して料金再計算
                    var withAmountFlags2 = _flagService?.Flags
                        .Where(f => f.Type == FlagType.WithAmount).ToList()
                        ?? new List<FlagDefinition>();

                    baseFee = rate.BaseFee;
                    foreach (var flag in withAmountFlags2)
                    {
                        if (!row.GetFlag(flag.Id)) continue;
                        if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                            baseFee = Math.Floor(rate.BaseFee * flag.AmountValue.Value);
                        else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                            baseFee = flag.AmountValue.Value;
                    }

                    if (row.D_YuryoKm.HasValue && row.D_YuryoKm > 0)
                    {
                        mileageFee = (Math.Floor((double)row.D_YuryoKm / 10) + 1) * rate.MileageFee;
                    }

                    if (isOotsuki && row.H_LateFeeOotsuki.HasValue)
                    {
                        lateFee = row.H_LateFeeOotsuki.Value;
                    }
                    else if (!isOotsuki && row.K_LateMinutes.HasValue && row.K_LateMinutes > 0)
                    {
                        double blocks = Math.Floor((double)row.K_LateMinutes / 30) + 1;
                        lateFee = rate.LateNightFixedFee + (blocks * rate.LateNightUnitFee);
                    }

                    totalFee = baseFee + mileageFee + lateFee;
                    totalRevenue += totalFee;
                }

                var transferRow = new TransferRowData
                {
                    RowIndex = row.RowIndex,
                    Day = row.B_Day?.ToString() ?? "-",
                    HansoCount = row.C_Hanso?.ToString() ?? "-",
                    YuryoKm = row.D_YuryoKm?.ToString("N0") ?? "-",
                    MuryoKm = row.E_MuryoKm?.ToString("N0") ?? "-",
                    IsKoryo = row.FlagSummaryText,
                    LateValue = isOotsuki
                        ? (row.H_LateFeeOotsuki?.ToString("N0") ?? "-")
                        : (row.K_LateMinutes?.ToString() ?? "-"),
                    BaseFee = baseFee > 0 ? baseFee.ToString("N0") : "-",
                    MileageFee = mileageFee > 0 ? mileageFee.ToString("N0") : "-",
                    LateFee = lateFee > 0 ? lateFee.ToString("N0") : "-",
                    TotalFee = totalFee > 0 ? totalFee.ToString("N0") : "-"
                };

                // バリデーション
                var issues = ValidateRow(row, SelectedVehicle.SheetName, isOotsuki);
                transferRow.HasError = issues.Any(i => i.Severity == IssueSeverity.Error);
                transferRow.HasWarning = issues.Any(i => i.Severity == IssueSeverity.Warning);

                foreach (var issue in issues)
                {
                    issue.Day = row.B_Day ?? 0;
                    _allValidationIssues.Add(issue);
                }

                CurrentVehicleRows.Add(transferRow);
            }

            // サマリー作成
            CurrentVehicleSummary = new VehicleSummary
            {
                WorkingDays = rows.Count(r => r.B_Day.HasValue),
                TotalHanso = totalHanso,
                TotalYuryoKm = totalYuryoKm,
                TotalMuryoKm = totalMuryoKm,
                TotalKm = totalYuryoKm + totalMuryoKm,
                KoryoCount = koryoCount,
                EstimatedRevenue = totalRevenue,
                AverageRevenuePerTrip = totalHanso > 0 ? totalRevenue / totalHanso : 0
            };

            FilterValidationIssues();
        }
    }
}
