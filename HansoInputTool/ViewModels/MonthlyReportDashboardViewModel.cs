using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using NLog;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 月報統計ダッシュボードのViewModel
    /// </summary>
    public class MonthlyReportDashboardViewModel : INotifyPropertyChanged
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly MonthlyReportScanner _scanner;
        private readonly MonthlyReportStatisticsService _statsService;

        private string _rootFolderPath;
        private List<MonthlyReportFile> _allReports;
        private ObservableCollection<MonthlyReportFile> _availableMonths;
        private MonthlyReportFile _selectedMonth;
        private Statistics _currentStatistics;
        private Statistics _yearlyStatistics;
        private bool _isLoading;
        private string _errorMessage;
        private bool _showYearly;
        private int _selectedYear;
        private List<int> _availableYears;

        public string RootFolderPath
        {
            get => _rootFolderPath;
            set
            {
                _rootFolderPath = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<MonthlyReportFile> AvailableMonths
        {
            get => _availableMonths;
            set
            {
                _availableMonths = value;
                OnPropertyChanged();
            }
        }

        public MonthlyReportFile SelectedMonth
        {
            get => _selectedMonth;
            set
            {
                _selectedMonth = value;
                OnPropertyChanged();
                if (value != null && !_showYearly)
                {
                    LoadMonthlyStatistics();
                }
            }
        }

        public Statistics CurrentStatistics
        {
            get => _currentStatistics;
            set
            {
                _currentStatistics = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasData));
            }
        }

        public Statistics YearlyStatistics
        {
            get => _yearlyStatistics;
            set
            {
                _yearlyStatistics = value;
                OnPropertyChanged();
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                _isLoading = value;
                OnPropertyChanged();
            }
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set
            {
                _errorMessage = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasError));
            }
        }

        public bool ShowYearly
        {
            get => _showYearly;
            set
            {
                _showYearly = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShowMonthly));

                if (value)
                {
                    LoadYearlyStatistics();
                }
                else if (SelectedMonth != null)
                {
                    LoadMonthlyStatistics();
                }
            }
        }

        public bool ShowMonthly => !ShowYearly;

        public int SelectedYear
        {
            get => _selectedYear;
            set
            {
                _selectedYear = value;
                OnPropertyChanged();
                LoadYearForSelection();
            }
        }

        public List<int> AvailableYears
        {
            get => _availableYears;
            set
            {
                _availableYears = value;
                OnPropertyChanged();
            }
        }

        public bool HasData => CurrentStatistics != null && CurrentStatistics.TotalHanso > 0;
        public bool HasError => !string.IsNullOrEmpty(ErrorMessage);
        public bool HasReports => _allReports != null && _allReports.Any();

        public string DisplayTitle => ShowYearly
            ? $"{SelectedYear}年度 年間統計"
            : SelectedMonth != null ? $"{SelectedMonth.DisplayName} 月次統計" : "統計ダッシュボード";

        public ICommand SelectFolderCommand { get; }
        public ICommand ExportReportCommand { get; }
        public ICommand ShowYearlyCommand { get; }
        public ICommand ShowMonthlyCommand { get; }

        public MonthlyReportDashboardViewModel()
        {
            _scanner = new MonthlyReportScanner();
            _statsService = new MonthlyReportStatisticsService();

            AvailableMonths = new ObservableCollection<MonthlyReportFile>();
            AvailableYears = new List<int>();

            SelectFolderCommand = new RelayCommand(_ => SelectFolder());
            ExportReportCommand = new RelayCommand(_ => ExportReport(), _ => HasData);
            ShowYearlyCommand = new RelayCommand(_ => ShowYearly = true);
            ShowMonthlyCommand = new RelayCommand(_ => ShowYearly = false);

            CurrentStatistics = new Statistics();
            YearlyStatistics = new Statistics();
        }

        /// <summary>
        /// フォルダ選択ダイアログを表示（WPF用）
        /// </summary>
        private void SelectFolder()
        {
            try
            {
                using (var dialog = new CommonOpenFileDialog())
                {
                    dialog.Title = "月報ファイルが格納されているフォルダを選択してください";
                    dialog.IsFolderPicker = true;
                    dialog.InitialDirectory = RootFolderPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    dialog.AddToMostRecentlyUsedList = false;
                    dialog.AllowNonFileSystemItems = false;
                    dialog.EnsureFileExists = true;
                    dialog.EnsurePathExists = true;
                    dialog.EnsureReadOnly = false;
                    dialog.EnsureValidNames = true;
                    dialog.Multiselect = false;
                    dialog.ShowPlacesList = true;

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        RootFolderPath = dialog.FileName;
                        ScanFolder();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "フォルダ選択エラー");
                ErrorMessage = $"フォルダを選択できませんでした: {ex.Message}";
            }
        }

        /// <summary>
        /// フォルダをスキャンして月報ファイルを検索
        /// </summary>
        private void ScanFolder()
        {
            try
            {
                IsLoading = true;
                ErrorMessage = string.Empty;

                Logger.Info($"フォルダスキャン開始: {RootFolderPath}");

                _allReports = _scanner.ScanMonthlyReports(RootFolderPath);

                if (!_allReports.Any())
                {
                    ErrorMessage = "月報ファイルが見つかりませんでした";
                    return;
                }

                // 年度リストを作成
                AvailableYears = _allReports
                    .Select(r => r.Year)
                    .Distinct()
                    .OrderByDescending(y => y)
                    .ToList();

                // 最新年度を選択
                if (AvailableYears.Any())
                {
                    _selectedYear = AvailableYears.First();
                    OnPropertyChanged(nameof(SelectedYear));
                    LoadYearForSelection();
                }

                Logger.Info($"月報ファイル検出完了: {_allReports.Count}件");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "フォルダスキャンエラー");
                ErrorMessage = $"フォルダのスキャンに失敗しました: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// 選択された年度の月報をロード
        /// </summary>
        private void LoadYearForSelection()
        {
            try
            {
                var yearReports = _scanner.GetReportsByYear(_allReports, SelectedYear);

                AvailableMonths.Clear();
                foreach (var report in yearReports)
                {
                    AvailableMonths.Add(report);
                }

                // 最新月を選択
                if (AvailableMonths.Any())
                {
                    SelectedMonth = AvailableMonths.Last();
                }

                // 年間統計を計算
                if (ShowYearly)
                {
                    LoadYearlyStatistics();
                }

                OnPropertyChanged(nameof(DisplayTitle));
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "年度データ読み込みエラー");
                ErrorMessage = $"年度データの読み込みに失敗しました: {ex.Message}";
            }
        }

        /// <summary>
        /// 月次統計を読み込み
        /// </summary>
        private void LoadMonthlyStatistics()
        {
            try
            {
                IsLoading = true;
                ErrorMessage = string.Empty;

                Logger.Info($"月次統計計算: {SelectedMonth.DisplayName}");

                CurrentStatistics = _statsService.CalculateStatisticsFromFile(SelectedMonth);

                OnPropertyChanged(nameof(DisplayTitle));

                Logger.Info("月次統計計算完了");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "月次統計計算エラー");
                ErrorMessage = $"統計の計算に失敗しました: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// 年間統計を読み込み
        /// </summary>
        private void LoadYearlyStatistics()
        {
            try
            {
                IsLoading = true;
                ErrorMessage = string.Empty;

                Logger.Info($"年間統計計算: {SelectedYear}年度");

                var yearReports = _scanner.GetReportsByYear(_allReports, SelectedYear);
                YearlyStatistics = _statsService.CalculateCombinedStatistics(yearReports);
                CurrentStatistics = YearlyStatistics;

                OnPropertyChanged(nameof(DisplayTitle));

                Logger.Info("年間統計計算完了");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "年間統計計算エラー");
                ErrorMessage = $"年間統計の計算に失敗しました: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// 統計レポートをエクスポート
        /// </summary>
        private void ExportReport()
        {
            try
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "テキストファイル (*.txt)|*.txt",
                    FileName = ShowYearly
                        ? $"年間統計レポート_{SelectedYear}年度_{DateTime.Now:yyyyMMdd}.txt"
                        : $"月次統計レポート_{SelectedMonth?.DisplayName?.Replace("年", "").Replace("月", "")}_{DateTime.Now:yyyyMMdd}.txt",
                    DefaultExt = ".txt"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var report = GenerateTextReport();
                    System.IO.File.WriteAllText(saveDialog.FileName, report, System.Text.Encoding.UTF8);

                    System.Windows.MessageBox.Show(
                        "統計レポートをエクスポートしました",
                        "完了",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "レポートエクスポートエラー");
                System.Windows.MessageBox.Show(
                    $"エクスポートに失敗しました: {ex.Message}",
                    "エラー",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// テキスト形式の統計レポートを生成
        /// </summary>
        private string GenerateTextReport()
        {
            var report = new System.Text.StringBuilder();
            report.AppendLine("=== 搬送統計レポート ===");
            report.AppendLine($"作成日時: {DateTime.Now:yyyy年MM月dd日 HH:mm:ss}");
            report.AppendLine($"対象: {DisplayTitle}");
            report.AppendLine($"データ元: {RootFolderPath}");
            report.AppendLine();

            report.AppendLine("【基本統計】");
            report.AppendLine($"総搬送回数: {CurrentStatistics.TotalHanso:N0} 回");
            report.AppendLine($"総有料キロ: {CurrentStatistics.TotalYuryoKm:N1} km");
            report.AppendLine($"総無料キロ: {CurrentStatistics.TotalMuryoKm:N1} km");
            report.AppendLine($"平均有料キロ: {CurrentStatistics.AverageYuryoKm:N1} km");
            report.AppendLine($"平均無料キロ: {CurrentStatistics.AverageMuryoKm:N1} km");
            report.AppendLine($"行旅回数: {CurrentStatistics.TotalKoryo} 回");
            report.AppendLine($"総深夜料金: ¥{CurrentStatistics.TotalLateCharges:N0}");
            report.AppendLine();

            report.AppendLine("【売上情報】");
            report.AppendLine($"推定売上: ¥{CurrentStatistics.EstimatedRevenue:N0}");
            report.AppendLine($"1回あたり平均売上: ¥{CurrentStatistics.AverageRevenuePerTrip:N0}");
            report.AppendLine();

            report.AppendLine("【車両情報】");
            report.AppendLine($"稼働車両数: {CurrentStatistics.ActiveVehicleCount} 台");
            report.AppendLine($"最多使用車両: {CurrentStatistics.MostUsedVehicle} ({CurrentStatistics.MostUsedVehicleCount}回)");
            report.AppendLine();

            report.AppendLine("【日別記録】");
            report.AppendLine($"営業日数: {CurrentStatistics.WorkingDays} 日");
            report.AppendLine($"1日平均搬送回数: {CurrentStatistics.AverageHansoPerDay:N1} 回");
            report.AppendLine($"1日最大搬送回数: {CurrentStatistics.MaxDailyHanso} 回 ({CurrentStatistics.MaxDailyHansoDate:M月d日})");
            report.AppendLine($"1日最大走行距離: {CurrentStatistics.MaxDailyKm:N1} km ({CurrentStatistics.MaxDailyKmDate:M月d日})");

            if (ShowYearly && AvailableMonths.Any())
            {
                report.AppendLine();
                report.AppendLine("【月別内訳】");
                foreach (var month in AvailableMonths)
                {
                    report.AppendLine($"- {month.DisplayName}: {month.FileName}");
                }
            }

            return report.ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}