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


// WPF型を明示（WindowsAPICodePack経由のSystem.Windows.Forms競合を解消）
using Control      = System.Windows.Controls.Control;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using TextBox      = System.Windows.Controls.TextBox;
using DataObject   = System.Windows.DataObject;
namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 通常シート（寝台車・霊柩車・CH系）の入力・バリデーション・登録を担当するViewModel
    /// </summary>
    public class NormalSheetViewModel : ObservableObject
    {
        private readonly ValidationService _validationService;
        private readonly InputValidator _inputValidator;

        private ExcelHandler _excelHandler;
        private Action<string> _log;
        private Action<int?> _updatePreview;
        private FlagDefinitionService _flagService;
        private VehicleSettingsService _vehicleSettingsService;

        // 入力対象の年・月を取得するデリゲート（月末日チェックに使用）
        private Func<(int year, int month)> _getYearMonth;

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
                    ClearValidationErrors();
                    _updatePreview?.Invoke(null);
                    OnPropertyChanged(nameof(IsOotsukiSheet));
                    OnPropertyChanged(nameof(IsFeeMode));
                    OnPropertyChanged(nameof(LateInputLabel));
                    OnPropertyChanged(nameof(IsFuelTrackedVehicle));
                    ClearValidationErrors();
                }
            }
        }

        public bool IsOotsukiSheet => IsFeeMode; // 後方互換のため残す
        public bool IsFeeMode => _vehicleSettingsService?.IsFeeMode(SelectedNormalSheet ?? "") ?? (SelectedNormalSheet?.Contains("大月") ?? false);
        public string LateInputLabel => IsFeeMode ? "深夜料金(H)" : "深夜時間(K)";

        /// <summary>選択中の車両が給油管理表への記録対象かどうか（設定画面でオン/オフ）</summary>
        public bool IsFuelTrackedVehicle => _vehicleSettingsService?.IsFuelTracked(SelectedNormalSheet ?? "") ?? false;

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

        #endregion

        #region 給油入力（給油管理対象車両のみ）

        private bool _isFuelChecked;
        /// <summary>「給油あり」チェックボックスの状態</summary>
        public bool IsFuelChecked
        {
            get => _isFuelChecked;
            set
            {
                if (SetProperty(ref _isFuelChecked, value))
                {
                    if (!value)
                    {
                        // チェックを外したら入力値とエラーもクリアする
                        FuelOdometerKm = string.Empty;
                        FuelLiters     = string.Empty;
                    }
                    ValidateInput();
                }
            }
        }

        private string _fuelOdometerKm;
        /// <summary>給油時Km</summary>
        public string FuelOdometerKm
        {
            get => _fuelOdometerKm;
            set { if (SetProperty(ref _fuelOdometerKm, value)) ValidateInput(); }
        }

        private string _fuelLiters;
        /// <summary>給油㍑数</summary>
        public string FuelLiters
        {
            get => _fuelLiters;
            set { if (SetProperty(ref _fuelLiters, value)) ValidateInput(); }
        }

        private string _fuelOdometerKmError;
        public string FuelOdometerKmError { get => _fuelOdometerKmError; set => SetProperty(ref _fuelOdometerKmError, value); }

        private string _fuelLitersError;
        public string FuelLitersError { get => _fuelLitersError; set => SetProperty(ref _fuelLitersError, value); }

        #endregion

        #region 動的フラグチェックボックス

        /// <summary>UIにバインドするフラグチェックボックスのリスト</summary>
        public ObservableCollection<FlagCheckBoxItem> FlagItems { get; } = new();

        /// <summary>フラグの状態を Id → bool で返す（登録時に使用）</summary>
        public Dictionary<string, bool> GetFlagStates()
            => FlagItems.ToDictionary(f => f.Id, f => f.IsChecked);

        /// <summary>フラグを全てリセットする</summary>
        public void ResetFlags()
        {
            foreach (var item in FlagItems) item.IsChecked = false;
        }

        /// <summary>指定IDのフラグをON/OFFトグルする（ショートカット用）</summary>
        public bool ToggleFlag(string flagId)
        {
            var item = FlagItems.FirstOrDefault(f => f.Id == flagId);
            if (item == null) return false;
            item.IsChecked = !item.IsChecked;
            return true;
        }

        /// <summary>フラグ定義が変更されたときにFlagItemsを再構築する</summary>
        public void RebuildFlagItems()
        {
            FlagItems.Clear();
            if (_flagService == null) return;
            foreach (var flag in _flagService.Flags.OrderBy(f => f.Order))
                FlagItems.Add(new FlagCheckBoxItem(flag));
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

        public void RefreshFeeMode()
        {
            OnPropertyChanged(nameof(IsFeeMode));
            OnPropertyChanged(nameof(IsOotsukiSheet));
            OnPropertyChanged(nameof(LateInputLabel));
        }

        public void Initialize(ExcelHandler excelHandler, Action<string> log, Action<int?> updatePreview, FlagDefinitionService flagService = null, VehicleSettingsService vehicleSettingsService = null, Func<(int year, int month)> getYearMonth = null)
        {
            _excelHandler           = excelHandler;
            _log                    = log;
            _updatePreview          = updatePreview;
            _flagService            = flagService;
            _vehicleSettingsService = vehicleSettingsService;
            _getYearMonth           = getYearMonth;
            RebuildFlagItems();
        }

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
                Day, YuryoKm, MuryoKm, LateValue, IsFeeMode, SelectedNormalSheet);

            DayError = result.DayError;
            YuryoKmError = result.YuryoKmError;
            MuryoKmError = result.MuryoKmError;
            LateValueError = result.LateValueError;

            bool fuelHasError = false;
            if (IsFuelChecked)
            {
                if (!double.TryParse(FuelOdometerKm, out var km) || km <= 0)
                {
                    FuelOdometerKmError = "給油時Kmを正しく入力してください。";
                    fuelHasError = true;
                }
                else FuelOdometerKmError = string.Empty;

                if (!double.TryParse(FuelLiters, out var l) || l <= 0)
                {
                    FuelLitersError = "給油㍑数を正しく入力してください。";
                    fuelHasError = true;
                }
                else FuelLitersError = string.Empty;
            }
            else
            {
                FuelOdometerKmError = string.Empty;
                FuelLitersError = string.Empty;
            }

            HasValidationErrors = result.HasErrors || fuelHasError;
        }

        public void ClearValidationErrors()
        {
            DayError = string.Empty;
            YuryoKmError = string.Empty;
            MuryoKmError = string.Empty;
            LateValueError = string.Empty;
            FuelOdometerKmError = string.Empty;
            FuelLitersError = string.Empty;
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

            if (IsFeeMode)
            {
                if (!TryParseValue(LateValue, "深夜料金(H)", out var lateVal)) return;
                values["深夜料金(H)"] = lateVal;
            }
            else
            {
                if (!TryParseValue(LateValue, "深夜時間(K)", out var lateVal)) return;
                values["深夜時間(K)"] = lateVal;
            }

            // 年・月を取得してバリデーションに渡す（月末日チェック用）
            var (valYear, valMonth) = _getYearMonth?.Invoke() ?? (0, 0);
            var validationResult = _validationService.ValidateNormalData(values, SelectedNormalSheet, valYear, valMonth);
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
                var flagStates = GetFlagStates();
                var (targetRow, insertInfo) = _excelHandler.RegisterNormalData(SelectedNormalSheet, values, flagStates);

                // 給油ありがチェックされていれば、続けて給油記録も登録する
                if (IsFuelChecked)
                {
                    double.TryParse(FuelOdometerKm, out var fuelKm);
                    double.TryParse(FuelLiters, out var fuelLiters);
                    _excelHandler.RegisterFuelData(SelectedNormalSheet, (int)dayVal.Value, fuelKm, fuelLiters);
                    _log?.Invoke($"[{SelectedNormalSheet}] {dayVal.Value}日の給油記録（{fuelKm:N0}km / {fuelLiters:N0}L）を登録しました。");
                }

                _updatePreview?.Invoke(targetRow);
                // DB使用時はSave()不要（DBへの書き込みは即時コミット済み）
                if (_excelHandler.DbService == null) _excelHandler.Save();
                if (!string.IsNullOrEmpty(insertInfo)) _log?.Invoke($"[{SelectedNormalSheet}] {insertInfo}");
                _log?.Invoke($"[{SelectedNormalSheet}] の {targetRow}行目にデータを登録しました。");

                Day = YuryoKm = MuryoKm = LateValue = string.Empty;
                IsFuelChecked = false;
                FuelOdometerKm = FuelLiters = string.Empty;
                ResetFlags();
                ClearValidationErrors();

                await Task.Delay(50);
                Messenger.Send(new FocusMessage { TargetElementName = "NormalDayTextBox" });
            }
            catch (InvalidOperationException ex)
            {
                // 確定済みセッションへの登録など、意図的にブロックしている操作
                _log?.Invoke($"登録ブロック: {ex.Message}");
                MessageBox.Show(ex.Message, "登録できません", MessageBoxButton.OK, MessageBoxImage.Warning);
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

    /// <summary>
    /// チェックボックス1件をXAMLにバインドするためのVM
    /// </summary>
    public class FlagCheckBoxItem : ObservableObject
    {
        public string Id          { get; }
        public string DisplayName { get; }
        public FlagType Type      { get; }

        private bool _isChecked;
        public bool IsChecked
        {
            get => _isChecked;
            set => SetProperty(ref _isChecked, value);
        }

        public FlagCheckBoxItem(FlagDefinition def)
        {
            Id          = def.Id;
            DisplayName = def.DisplayName;
            Type        = def.Type;
        }
    }
}


