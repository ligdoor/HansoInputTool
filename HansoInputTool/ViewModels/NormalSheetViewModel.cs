using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Messaging;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 通常シート（寝台車・霊柩車・CH系）の入力・バリデーション・登録を担当するViewModel
    /// </summary>
    public class NormalSheetViewModel : ObservableObject
    {
        private readonly ValidationService _validationService;
        private readonly InputValidator _inputValidator;

        // ExcelHandler と PreviewData は MainViewModel から注入する
        private ExcelHandler _excelHandler;
        private Action<string> _log;
        private Action _updatePreview;

        #region シート選択

        public ObservableCollection<string> NormalSheets { get; } = new();

        private string _selectedNormalSheet;
        public string SelectedNormalSheet
        {
            get => _selectedNormalSheet;
            set
            {
                if (SetProperty(ref _selectedNormalSheet, value))
                {
                    _updatePreview?.Invoke();
                    OnPropertyChanged(nameof(IsOotsukiSheet));
                    ClearValidationErrors();
                }
            }
        }

        public bool IsOotsukiSheet => SelectedNormalSheet?.Contains("大月") ?? false;

        #endregion

        #region 入力フィールド

        private string _day;
        public string Day
        {
            get => _day;
            set { if (SetProperty(ref _day, value)) ValidateInput(); }
        }

        private string _yuryoKm;
        public string YuryoKm
        {
            get => _yuryoKm;
            set { if (SetProperty(ref _yuryoKm, value)) ValidateInput(); }
        }

        private string _muryoKm;
        public string MuryoKm
        {
            get => _muryoKm;
            set { if (SetProperty(ref _muryoKm, value)) ValidateInput(); }
        }

        private string _lateValue;
        public string LateValue
        {
            get => _lateValue;
            set { if (SetProperty(ref _lateValue, value)) ValidateInput(); }
        }

        private bool _isKoryo;
        public bool IsKoryo
        {
            get => _isKoryo;
            set => SetProperty(ref _isKoryo, value);
        }

        private bool _isEmbalming;
        public bool IsEmbalming
        {
            get => _isEmbalming;
            set => SetProperty(ref _isEmbalming, value);
        }

        #endregion

        #region バリデーションエラー

        private string _dayError;
        public string DayError { get => _dayError; set => SetProperty(ref _dayError, value); }

        private string _yuryoKmError;
        public string YuryoKmError { get => _yuryoKmError; set => SetProperty(ref _yuryoKmError, value); }

        private string _muryoKmError;
        public string MuryoKmError { get => _muryoKmError; set => SetProperty(ref _muryoKmError, value); }

        private string _lateValueError;
        public string LateValueError { get => _lateValueError; set => SetProperty(ref _lateValueError, value); }

        private bool _hasValidationErrors;
        public bool HasValidationErrors
        {
            get => _hasValidationErrors;
            set
            {
                if (SetProperty(ref _hasValidationErrors, value))
                    CommandManager.InvalidateRequerySuggested();
            }
        }

        #endregion

        #region コマンド

        public ICommand RegisterCommand { get; }
        // XAMLバインディング互換用エイリアス
        public ICommand RegisterNormalCommand => RegisterCommand;

        #endregion

        public NormalSheetViewModel(ValidationService validationService)
        {
            _validationService = validationService;
            _inputValidator = new InputValidator(validationService);
            RegisterCommand = new RelayCommand(
                async p => await RegisterAsync(),
                p => !HasValidationErrors);
        }

        /// <summary>
        /// ExcelHandler・ログ・プレビュー更新を MainViewModel から受け取る
        /// </summary>
        public void Initialize(ExcelHandler excelHandler, Action<string> log, Action updatePreview)
        {
            _excelHandler = excelHandler;
            _log = log;
            _updatePreview = updatePreview;
        }

        /// <summary>
        /// シートリストを再構築する
        /// </summary>
        public void PopulateSheets(List<string> allVehicleSheets, string previousSelection)
        {
            NormalSheets.Clear();
            foreach (var s in allVehicleSheets.Where(s => !s.Contains("東日本")))
                NormalSheets.Add(s);

            SelectedNormalSheet = NormalSheets.Contains(previousSelection)
                ? previousSelection
                : NormalSheets.FirstOrDefault();
        }

        #region バリデーション

        private void ValidateInput()
        {
            var result = _inputValidator.ValidateNormalSheet(
                Day, YuryoKm, MuryoKm, LateValue, IsOotsukiSheet, SelectedNormalSheet);

            DayError = result.DayError;
            YuryoKmError = result.YuryoKmError;
            MuryoKmError = result.MuryoKmError;
            LateValueError = result.LateValueError;
            HasValidationErrors = result.HasErrors;
        }

        public void ClearValidationErrors()
        {
            DayError = string.Empty;
            YuryoKmError = string.Empty;
            MuryoKmError = string.Empty;
            LateValueError = string.Empty;
            HasValidationErrors = false;
        }

        #endregion

        #region 登録処理

        private async Task RegisterAsync()
        {
            if (string.IsNullOrEmpty(SelectedNormalSheet))
            {
                MessageBox.Show("通常シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(Day))
            {
                MessageBox.Show("日付は必須です。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var values = new Dictionary<string, double?>();
            if (!TryParseValue(Day, "日(B)", out var dayVal)) return;
            values["日(B)"] = dayVal;

            if (!TryParseValue(YuryoKm, "有料キロ(D)", out var yuryoKmVal)) return;
            values["有料キロ(D)"] = yuryoKmVal.HasValue ? Math.Round(yuryoKmVal.Value, MidpointRounding.AwayFromZero) : null;

            if (!TryParseValue(MuryoKm, "無料キロ(E)", out var muryoKmVal)) return;
            values["無料キロ(E)"] = muryoKmVal.HasValue ? Math.Round(muryoKmVal.Value, MidpointRounding.AwayFromZero) : null;

            if (IsOotsukiSheet)
            {
                if (!TryParseValue(LateValue, "深夜料金(H)", out var lateVal)) return;
                values["深夜料金(H)"] = lateVal;
            }
            else
            {
                if (!TryParseValue(LateValue, "深夜時間(K)", out var lateVal)) return;
                values["深夜時間(K)"] = lateVal;
            }

            var validationResult = _validationService.ValidateNormalData(values, SelectedNormalSheet);
            if (!validationResult.IsValid)
            {
                MessageBox.Show($"入力内容にエラーがあります:\n\n{validationResult.GetErrorMessage()}",
                    "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (validationResult.HasWarnings)
            {
                var confirm = MessageBox.Show(
                    $"以下の警告があります:\n\n{validationResult.GetWarningMessage()}\n\nそのまま登録しますか？",
                    "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;
            }

            try
            {
                var (targetRow, insertInfo) = _excelHandler.RegisterNormalData(SelectedNormalSheet, values, IsKoryo, IsEmbalming);
                _updatePreview?.Invoke();
                _excelHandler.Save();
                if (!string.IsNullOrEmpty(insertInfo)) _log?.Invoke($"[{SelectedNormalSheet}] {insertInfo}");
                _log?.Invoke($"[{SelectedNormalSheet}] の {targetRow}行目にデータを登録しました。");

                Day = YuryoKm = MuryoKm = LateValue = string.Empty;
                IsKoryo = false;
                IsEmbalming = false;
                ClearValidationErrors();

                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "NormalDayTextBox" });
            }
            catch (Exception ex)
            {
                _log?.Invoke($"登録エラー: {ex.Message}");
                MessageBox.Show("登録エラーが発生しました。\n詳細はログファイルを確認してください。",
                    "登録エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        private static bool TryParseValue(string input, string fieldName, out double? result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input)) return true;
            if (double.TryParse(input, out double parsed)) { result = parsed; return true; }
            MessageBox.Show($"「{input}」は {fieldName} の数値として認識できません。",
                "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }
}
