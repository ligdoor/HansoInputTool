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
    public class TransferConfirmationViewModel : ObservableObject
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly ExcelHandler _excelHandler;
        private readonly Dictionary<string, RateInfo> _rates;
        private readonly ColumnMapping _columnMap;
        private readonly Action<bool> _callback;
        private readonly FlagDefinitionService _flagService;

        // 基本情報
        public string Period { get; set; }
        public string Month { get; set; }
        public string RNumber { get; set; }

        // 車両リスト
        public ObservableCollection<VehicleTab> VehicleTabs { get; set; }

        private VehicleTab _selectedVehicle;
        public VehicleTab SelectedVehicle
        {
            get => _selectedVehicle;
            set
            {
                if (SetProperty(ref _selectedVehicle, value))
                {
                    LoadVehicleData();
                }
            }
        }

        // 現在の車両データ
        public ObservableCollection<TransferRowData> CurrentVehicleRows { get; set; }

        private VehicleSummary _currentVehicleSummary;
        public VehicleSummary CurrentVehicleSummary
        {
            get => _currentVehicleSummary;
            set => SetProperty(ref _currentVehicleSummary, value);
        }

        // エラー・警告リスト
        public ObservableCollection<ValidationIssue> ValidationIssues { get; set; }
        private List<ValidationIssue> _allValidationIssues = new List<ValidationIssue>();

        private bool _showErrorsOnly;
        public bool ShowErrorsOnly
        {
            get => _showErrorsOnly;
            set
            {
                if (SetProperty(ref _showErrorsOnly, value))
                {
                    FilterValidationIssues();
                }
            }
        }

        // 統計情報
        public string TotalVehicles { get; set; }
        public string VehiclesWithErrors { get; set; }
        public string VehiclesWithWarnings { get; set; }
        public string TotalEstimatedRevenue { get; set; }

        // コマンド
        public ICommand PreviousVehicleCommand { get; }
        public ICommand NextVehicleCommand { get; }
        public ICommand EditDataCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand ConfirmTransferCommand { get; }
        public ICommand JumpToIssueCommand { get; }

        public TransferConfirmationViewModel(
            ExcelHandler excelHandler,
            Dictionary<string, RateInfo> rates,
            ColumnMapping columnMap,
            string period,
            string month,
            string rNumber,
            Action<bool> callback,
            FlagDefinitionService flagService = null)
        {
            _excelHandler = excelHandler;
            _rates        = rates;
            _columnMap    = columnMap;
            _callback     = callback;
            _flagService  = flagService;

            Period = period;
            Month = month;
            RNumber = rNumber;

            VehicleTabs = new ObservableCollection<VehicleTab>();
            CurrentVehicleRows = new ObservableCollection<TransferRowData>();
            ValidationIssues = new ObservableCollection<ValidationIssue>();

            // コマンド初期化
            PreviousVehicleCommand = new RelayCommand(_ => MoveToPreviousVehicle(), _ => CanMoveToPrevious());
            NextVehicleCommand = new RelayCommand(_ => MoveToNextVehicle(), _ => CanMoveToNext());
            EditDataCommand = new RelayCommand(_ => EditData());
            CancelCommand = new RelayCommand(_ => Cancel());
            ConfirmTransferCommand = new RelayCommand(_ => ConfirmTransfer());
            JumpToIssueCommand = new RelayCommand(param => JumpToIssue(param as ValidationIssue));

            // データ読み込み
            LoadAllVehicles();
        }

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
            bool isOotsuki = sheetName.Contains("大月");
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
            bool isOotsuki = SelectedVehicle.SheetName.Contains("大月");
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

        private SolidColorBrush GetStatusColor(int errors, int warnings)
        {
            if (errors > 0)
                return new SolidColorBrush(Color.FromRgb(239, 68, 68));
            if (warnings > 0)
                return new SolidColorBrush(Color.FromRgb(251, 191, 36));
            return new SolidColorBrush(Color.FromRgb(16, 185, 129));
        }

        private string GetStatusIcon(int errors, int warnings)
        {
            if (errors > 0) return "❌";
            if (warnings > 0) return "⚠️";
            return "✓";
        }

        private void MoveToPreviousVehicle()
        {
            var index = VehicleTabs.IndexOf(SelectedVehicle);
            if (index > 0)
            {
                SelectedVehicle = VehicleTabs[index - 1];
            }
        }

        private void MoveToNextVehicle()
        {
            var index = VehicleTabs.IndexOf(SelectedVehicle);
            if (index < VehicleTabs.Count - 1)
            {
                SelectedVehicle = VehicleTabs[index + 1];
            }
        }

        private bool CanMoveToPrevious()
        {
            return SelectedVehicle != null && VehicleTabs.IndexOf(SelectedVehicle) > 0;
        }

        private bool CanMoveToNext()
        {
            return SelectedVehicle != null && VehicleTabs.IndexOf(SelectedVehicle) < VehicleTabs.Count - 1;
        }

        private void JumpToIssue(ValidationIssue issue)
        {
            if (issue == null) return;

            var vehicle = VehicleTabs.FirstOrDefault(v => v.SheetName == issue.SheetName);
            if (vehicle != null)
            {
                SelectedVehicle = vehicle;
            }
        }

        private void EditData()
        {
            _callback?.Invoke(false);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }

        private void Cancel()
        {
            _callback?.Invoke(false);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }

        private void ConfirmTransfer()
        {
            var totalErrors = VehicleTabs.Sum(v => v.ErrorCount);
            if (totalErrors > 0)
            {
                var result = MessageBox.Show(
                    $"{totalErrors}件のエラーが検出されています。\nこのまま転記を実行しますか？",
                    "エラー確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            _callback?.Invoke(true);
            Application.Current.Windows.OfType<TransferConfirmationWindow>().FirstOrDefault()?.Close();
        }
    }

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