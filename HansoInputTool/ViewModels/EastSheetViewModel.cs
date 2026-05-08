using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Messaging;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 東日本シートの入力・バリデーション・登録を担当するViewModel
    /// </summary>
    public class EastSheetViewModel : ObservableObject
    {
        private readonly ValidationService _validationService;
        private readonly InputValidator _inputValidator;

        private ExcelHandler _excelHandler;
        private Action<string> _log;

        #region シート選択

        public ObservableCollection<string> EastSheets { get; } = new();
        private readonly List<string> _registeredSheets = new();

        private string _selectedEastSheet;
        public string SelectedEastSheet
        {
            get => _selectedEastSheet;
            set
            {
                if (SetProperty(ref _selectedEastSheet, value))
                {
                    LoadExistingValues();
                    UpdateSheetStatus();
                    ClearValidationErrors();
                }
            }
        }

        private string _sheetStatus = "（未登録）";
        public string SheetStatus { get => _sheetStatus; set => SetProperty(ref _sheetStatus, value); }

        private bool _isSheetRegistered;
        public bool IsSheetRegistered { get => _isSheetRegistered; set => SetProperty(ref _isSheetRegistered, value); }

        #endregion

        #region 入力フィールド

        private string _jitsudo;
        public string Jitsudo
        {
            get => _jitsudo;
            set { if (SetProperty(ref _jitsudo, value)) ValidateInput(); }
        }

        private string _hanso;
        public string Hanso
        {
            get => _hanso;
            set { if (SetProperty(ref _hanso, value)) ValidateInput(); }
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

        private string _unso;
        public string Unso
        {
            get => _unso;
            set { if (SetProperty(ref _unso, value)) ValidateInput(); }
        }

        #endregion

        #region バリデーションエラー

        private string _jitsudoError;
        public string JitsudoError { get => _jitsudoError; set => SetProperty(ref _jitsudoError, value); }

        private string _hansoError;
        public string HansoError { get => _hansoError; set => SetProperty(ref _hansoError, value); }

        private string _yuryoKmError;
        public string YuryoKmError { get => _yuryoKmError; set => SetProperty(ref _yuryoKmError, value); }

        private string _muryoKmError;
        public string MuryoKmError { get => _muryoKmError; set => SetProperty(ref _muryoKmError, value); }

        private string _unsoError;
        public string UnsoError { get => _unsoError; set => SetProperty(ref _unsoError, value); }

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
        public ICommand RegisterEastCommand => RegisterCommand;

        #endregion

        public EastSheetViewModel(ValidationService validationService)
        {
            _validationService = validationService;
            _inputValidator = new InputValidator(validationService);
            RegisterCommand = new RelayCommand(
                async p => await RegisterAsync(),
                p => !HasValidationErrors);
        }

        /// <summary>
        /// ExcelHandler・ログを MainViewModel から受け取る
        /// </summary>
        public void Initialize(ExcelHandler excelHandler, Action<string> log)
        {
            _excelHandler = excelHandler;
            _log = log;
        }

        /// <summary>
        /// シートリストを再構築する
        /// </summary>
        public void PopulateSheets(List<string> allVehicleSheets, string previousSelection)
        {
            EastSheets.Clear();
            foreach (var s in allVehicleSheets.Where(s => s.Contains("東日本")))
                EastSheets.Add(s);

            SelectedEastSheet = EastSheets.Contains(previousSelection)
                ? previousSelection
                : EastSheets.FirstOrDefault();
        }

        /// <summary>
        /// 転記完了後など、登録済みリストをリセットしフィールドもクリアする
        /// </summary>
        public void ClearRegisteredSheets()
        {
            _registeredSheets.Clear();
            ClearFields();
            UpdateSheetStatus();
        }

        #region バリデーション

        private void ValidateInput()
        {
            var result = _inputValidator.ValidateEastSheet(Jitsudo, Hanso, YuryoKm, MuryoKm, Unso);

            JitsudoError = result.JitsudoError;
            HansoError = result.HansoError;
            YuryoKmError = result.YuryoKmError;
            MuryoKmError = result.MuryoKmError;
            UnsoError = result.UnsoError;
            HasValidationErrors = result.HasErrors;
        }

        public void ClearValidationErrors()
        {
            JitsudoError = string.Empty;
            HansoError = string.Empty;
            YuryoKmError = string.Empty;
            MuryoKmError = string.Empty;
            UnsoError = string.Empty;
            HasValidationErrors = false;
        }

        private void UpdateSheetStatus()
        {
            if (string.IsNullOrEmpty(SelectedEastSheet)) { IsSheetRegistered = false; SheetStatus = ""; return; }

            // Excelに実際に値が入っているかで登録済みを判定
            bool hasData = HasExistingData();
            if (hasData || _registeredSheets.Contains(SelectedEastSheet))
            {
                IsSheetRegistered = true;
                SheetStatus = "✅ 登録済み";
            }
            else
            {
                IsSheetRegistered = false;
                SheetStatus = "（未登録）";
            }
        }

        private bool HasExistingData()
        {
            if (_excelHandler == null || string.IsNullOrEmpty(SelectedEastSheet)) return false;
            var vals = _excelHandler.GetEastSheetValues(SelectedEastSheet);
            if (vals == null) return false;
            return vals.Values.Any(v => v != null);
        }

        /// <summary>
        /// シート選択時にExcelの登録済み値をテキストボックスに読み戻す
        /// </summary>
        private void LoadExistingValues()
        {
            if (_excelHandler == null || string.IsNullOrEmpty(SelectedEastSheet))
            {
                ClearFields();
                return;
            }

            var vals = _excelHandler.GetEastSheetValues(SelectedEastSheet);
            if (vals == null)
            {
                ClearFields();
                return;
            }

            // バリデーションが走らないよう、バッキングフィールドに直接セット後にまとめてPropertyChanged
            _jitsudo  = vals["延実働車輌数"]?.ToString() ?? string.Empty;
            _hanso    = vals["搬送回数"]?.ToString() ?? string.Empty;
            _yuryoKm  = vals["有料キロ数"]?.ToString() ?? string.Empty;
            _muryoKm  = vals["無料キロ数"]?.ToString() ?? string.Empty;
            _unso     = vals["運輸実績"]?.ToString() ?? string.Empty;

            OnPropertyChanged(nameof(Jitsudo));
            OnPropertyChanged(nameof(Hanso));
            OnPropertyChanged(nameof(YuryoKm));
            OnPropertyChanged(nameof(MuryoKm));
            OnPropertyChanged(nameof(Unso));
        }

        private void ClearFields()
        {
            _jitsudo = _hanso = _yuryoKm = _muryoKm = _unso = string.Empty;
            OnPropertyChanged(nameof(Jitsudo));
            OnPropertyChanged(nameof(Hanso));
            OnPropertyChanged(nameof(YuryoKm));
            OnPropertyChanged(nameof(MuryoKm));
            OnPropertyChanged(nameof(Unso));
        }

        #endregion

        #region 登録処理

        private async Task RegisterAsync()
        {
            if (string.IsNullOrEmpty(SelectedEastSheet))
            {
                MessageBox.Show("東日本シートが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var values = new Dictionary<string, double?>();
            if (!TryParseValue(Jitsudo, "延実働車輌数", out var jitsudo)) return; values["延実働車輌数"] = jitsudo;
            if (!TryParseValue(Hanso, "搬送回数", out var hanso)) return; values["搬送回数"] = hanso;
            if (!TryParseValue(YuryoKm, "有料キロ数", out var yuryo)) return; values["有料キロ数"] = yuryo;
            if (!TryParseValue(MuryoKm, "無料キロ数", out var muryo)) return; values["無料キロ数"] = muryo;
            if (!TryParseValue(Unso, "運輸実績", out var unso)) return; values["運輸実績"] = unso;

            var validationResult = _validationService.ValidateEastData(values);
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
                _excelHandler.RegisterEastData(SelectedEastSheet, values);
                _excelHandler.Save();
                _log?.Invoke($"[{SelectedEastSheet}] のデータを登録しました。");

                if (!_registeredSheets.Contains(SelectedEastSheet))
                    _registeredSheets.Add(SelectedEastSheet);
                UpdateSheetStatus();

                ClearValidationErrors();

                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "EastJitsudoTextBox" });
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
