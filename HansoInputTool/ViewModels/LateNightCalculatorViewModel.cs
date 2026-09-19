using System;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 会社出発時刻・故人宅出発時刻を入力し、深夜割増（22:00〜翌5:00）に該当する時間と
    /// 深夜料金を計算するポップアップウィンドウ用のViewModel。
    /// </summary>
    public class LateNightCalculatorViewModel : ObservableObject
    {
        private readonly RateInfo _rate;

        /// <summary>true: 深夜料金(円)を確定値として扱うシート（例: CH大月）。false: 深夜時間(分)を扱う通常シート。</summary>
        public bool IsFeeMode { get; }

        /// <summary>呼び出し元に反映するための表示ラベル（"深夜料金(H)" または "深夜時間(K)"）</summary>
        public string TargetLabel => IsFeeMode ? "深夜料金(H)" : "深夜時間(K)";

        private string _companyHour = string.Empty;
        /// <summary>会社出発時刻・時（0〜23）</summary>
        public string CompanyHour
        {
            get => _companyHour;
            set { if (SetProperty(ref _companyHour, value)) { ClearResult(); } }
        }

        private string _companyMinute = string.Empty;
        /// <summary>会社出発時刻・分（0〜59）</summary>
        public string CompanyMinute
        {
            get => _companyMinute;
            set { if (SetProperty(ref _companyMinute, value)) { ClearResult(); } }
        }

        private string _homeHour = string.Empty;
        /// <summary>故人宅出発時刻・時（0〜23）</summary>
        public string HomeHour
        {
            get => _homeHour;
            set { if (SetProperty(ref _homeHour, value)) { ClearResult(); } }
        }

        private string _homeMinute = string.Empty;
        /// <summary>故人宅出発時刻・分（0〜59）</summary>
        public string HomeMinute
        {
            get => _homeMinute;
            set { if (SetProperty(ref _homeMinute, value)) { ClearResult(); } }
        }

        private string _errorMessage;
        public string ErrorMessage { get => _errorMessage; set => SetProperty(ref _errorMessage, value); }

        private bool _hasResult;
        public bool HasResult { get => _hasResult; set { if (SetProperty(ref _hasResult, value)) CommandManager.InvalidateRequerySuggested(); } }

        private int _resultMinutes;
        public int ResultMinutes { get => _resultMinutes; set => SetProperty(ref _resultMinutes, value); }

        private string _resultSummaryText;
        /// <summary>計算結果の説明文（例："90分が深夜時間帯に該当します"）</summary>
        public string ResultSummaryText { get => _resultSummaryText; set => SetProperty(ref _resultSummaryText, value); }

        private string _resultFeeText;
        /// <summary>深夜料金の表示文字列（料金モードのシートのみ使用）</summary>
        public string ResultFeeText { get => _resultFeeText; set => SetProperty(ref _resultFeeText, value); }

        /// <summary>反映確定後、呼び出し元（NormalSheetViewModel.LateValue）に書き込む値。分または円の数字文字列。</summary>
        public string AppliedValue { get; private set; }

        public ICommand CalculateCommand { get; }
        public ICommand ApplyCommand { get; }
        public ICommand CancelCommand { get; }

        public LateNightCalculatorViewModel(bool isFeeMode, RateInfo rate)
        {
            IsFeeMode = isFeeMode;
            _rate = rate;

            CalculateCommand = new RelayCommand(_ => Calculate());
            ApplyCommand = new RelayCommand(p => Apply(p), _ => HasResult);
            CancelCommand = new RelayCommand(p => Cancel(p));
        }

        private void ClearResult()
        {
            HasResult = false;
            ErrorMessage = null;
        }

        private void Calculate()
        {
            ErrorMessage = null;
            HasResult = false;

            if (!TryParseTime(CompanyHour, CompanyMinute, out var start))
            {
                ErrorMessage = "会社出発時間を「時」0〜23・「分」0〜59の範囲で入力してください。";
                return;
            }
            if (!TryParseTime(HomeHour, HomeMinute, out var end))
            {
                ErrorMessage = "故人宅出発時間を「時」0〜23・「分」0〜59の範囲で入力してください。";
                return;
            }

            int minutes = LateNightFeeCalculator.CalculateNightMinutes(start, end);
            ResultMinutes = minutes;

            if (minutes <= 0)
            {
                ResultSummaryText = "深夜時間帯（22:00〜翌5:00）に該当する時間はありません。";
                ResultFeeText = null;
            }
            else if (IsFeeMode)
            {
                int fee = LateNightFeeCalculator.CalculateFee(minutes, _rate);
                ResultSummaryText = $"深夜時間帯に {minutes}分 該当します。";
                ResultFeeText = _rate == null
                    ? "（料金表が見つからないため金額は計算できません）"
                    : $"深夜料金：{fee:N0}円";
            }
            else
            {
                ResultSummaryText = $"深夜時間帯に {minutes}分 該当します。";
                ResultFeeText = null;
            }

            HasResult = true;
        }

        private static bool TryParseTime(string hourText, string minuteText, out TimeSpan time)
        {
            time = default;
            if (!int.TryParse(hourText, NumberStyles.None, CultureInfo.InvariantCulture, out int hour)) return false;
            if (!int.TryParse(minuteText, NumberStyles.None, CultureInfo.InvariantCulture, out int minute)) return false;
            if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return false;

            time = new TimeSpan(hour, minute, 0);
            return true;
        }

        private void Apply(object parameter)
        {
            if (!HasResult) return;

            if (IsFeeMode)
            {
                int fee = LateNightFeeCalculator.CalculateFee(ResultMinutes, _rate);
                AppliedValue = fee.ToString();
            }
            else
            {
                AppliedValue = ResultMinutes.ToString();
            }

            if (parameter is Window window)
            {
                window.DialogResult = true;
                window.Close();
            }
        }

        private void Cancel(object parameter)
        {
            if (parameter is Window window)
            {
                window.DialogResult = false;
                window.Close();
            }
        }
    }
}
