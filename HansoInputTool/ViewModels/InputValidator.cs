using System;
using System.Collections.Generic;
using HansoInputTool.Services;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// 入力値のリアルタイムバリデーションを担当するクラス
    /// MainViewModelから分離して、バリデーションロジックを独立管理
    /// </summary>
    public class InputValidator
    {
        private const string WarningPrefix = "⚠ ";
        private const string ParseErrorMessage = "数値を入力してください";

        private readonly ValidationService _validationService;

        public InputValidator(ValidationService validationService)
        {
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
        }

        #region Normal Sheet Validation

        /// <summary>
        /// 通常シートの入力値をバリデーション
        /// </summary>
        public NormalSheetValidationResult ValidateNormalSheet(
            string day,
            string yuryoKm,
            string muryoKm,
            string lateValue,
            bool isOotsukiSheet,
            string selectedSheet)
        {
            var result = new NormalSheetValidationResult();
            var values = new Dictionary<string, double?>();
            bool hasParseError = false;

            // 各入力値のパース処理
            hasParseError |= !ValidateAndParseField(day, "日(B)", val => result.DayError = val, values, shouldRound: false);
            hasParseError |= !ValidateAndParseField(yuryoKm, "有料キロ(D)", val => result.YuryoKmError = val, values, shouldRound: true);
            hasParseError |= !ValidateAndParseField(muryoKm, "無料キロ(E)", val => result.MuryoKmError = val, values, shouldRound: true);

            // 深夜時間/料金（シートによって項目名が変わる）
            if (!string.IsNullOrWhiteSpace(lateValue))
            {
                string fieldName = isOotsukiSheet ? "深夜料金(H)" : "深夜時間(K)";
                if (!TryParseValue(lateValue, out var lateVal))
                {
                    result.LateValueError = ParseErrorMessage;
                    hasParseError = true;
                }
                else
                {
                    values[fieldName] = lateVal;
                }
            }

            // パースエラーがある場合は、ビジネスルールバリデーションをスキップ
            if (hasParseError)
            {
                result.HasErrors = true;
                return result;
            }

            // ビジネスルールバリデーション
            if (values.Count > 0)
            {
                var validationResult = _validationService.ValidateNormalData(values, selectedSheet ?? "");
                ApplyNormalValidationResult(validationResult, result);
                result.HasErrors = !validationResult.IsValid;
            }

            return result;
        }

        /// <summary>
        /// 通常シートの単一フィールドのパースとバリデーション
        /// </summary>
        private bool ValidateAndParseField(
            string input,
            string fieldName,
            Action<string> setError,
            Dictionary<string, double?> values,
            bool shouldRound)
        {
            if (string.IsNullOrWhiteSpace(input))
                return true;

            if (!TryParseValue(input, out var parsedValue))
            {
                setError(ParseErrorMessage);
                return false;
            }

            values[fieldName] = shouldRound && parsedValue.HasValue
                ? Math.Round(parsedValue.Value, MidpointRounding.AwayFromZero)
                : parsedValue;
            return true;
        }

        /// <summary>
        /// 通常シートのバリデーション結果をResultオブジェクトに反映
        /// </summary>
        private void ApplyNormalValidationResult(ValidationResult validationResult, NormalSheetValidationResult result)
        {
            // エラーメッセージの設定
            SetErrorIfExists(validationResult.Errors, "日付", val => result.DayError = val);
            SetErrorIfExists(validationResult.Errors, "有料キロ", val => result.YuryoKmError = val);
            SetErrorIfExists(validationResult.Errors, "無料キロ", val => result.MuryoKmError = val);
            SetErrorIfExists(validationResult.Errors, "深夜時間", val => result.LateValueError = val);
            SetErrorIfExists(validationResult.Errors, "深夜料金", val => result.LateValueError = val);

            // 警告メッセージの設定（エラーがない場合のみ）
            SetWarningIfExists(validationResult.Warnings, "有料キロ", () => result.YuryoKmError, val => result.YuryoKmError = val);
            SetWarningIfExists(validationResult.Warnings, "無料キロ", () => result.MuryoKmError, val => result.MuryoKmError = val);
            SetWarningIfExists(validationResult.Warnings, "深夜時間", () => result.LateValueError, val => result.LateValueError = val);
            SetWarningIfExists(validationResult.Warnings, "深夜料金", () => result.LateValueError, val => result.LateValueError = val);

            // 走行距離の警告（有料・無料キロの両方にエラーがない場合）
            if (validationResult.Warnings.ContainsKey("走行距離") &&
                string.IsNullOrEmpty(result.YuryoKmError) &&
                string.IsNullOrEmpty(result.MuryoKmError))
            {
                result.YuryoKmError = WarningPrefix + string.Join(", ", validationResult.Warnings["走行距離"]);
            }
        }

        #endregion

        #region East Sheet Validation

        /// <summary>
        /// 東日本シートの入力値をバリデーション
        /// </summary>
        public EastSheetValidationResult ValidateEastSheet(
            string jitsudo,
            string hanso,
            string yuryoKm,
            string muryoKm,
            string unso)
        {
            var result = new EastSheetValidationResult();
            var values = new Dictionary<string, double?>();
            bool hasParseError = false;

            // 各入力値のパース処理
            hasParseError |= !ValidateAndParseEastField(jitsudo, "延実働車輌数", val => result.JitsudoError = val, values);
            hasParseError |= !ValidateAndParseEastField(hanso, "搬送回数", val => result.HansoError = val, values);
            hasParseError |= !ValidateAndParseEastField(yuryoKm, "有料キロ数", val => result.YuryoKmError = val, values);
            hasParseError |= !ValidateAndParseEastField(muryoKm, "無料キロ数", val => result.MuryoKmError = val, values);
            hasParseError |= !ValidateAndParseEastField(unso, "運輸実績", val => result.UnsoError = val, values);

            // パースエラーがある場合は、ビジネスルールバリデーションをスキップ
            if (hasParseError)
            {
                result.HasErrors = true;
                return result;
            }

            // ビジネスルールバリデーション
            if (values.Count > 0)
            {
                var validationResult = _validationService.ValidateEastData(values);
                ApplyEastValidationResult(validationResult, result);
                result.HasErrors = !validationResult.IsValid;
            }

            return result;
        }

        /// <summary>
        /// 東日本シートの単一フィールドのパースとバリデーション
        /// </summary>
        private bool ValidateAndParseEastField(
            string input,
            string fieldName,
            Action<string> setError,
            Dictionary<string, double?> values)
        {
            if (string.IsNullOrWhiteSpace(input))
                return true;

            if (!TryParseValue(input, out var parsedValue))
            {
                setError(ParseErrorMessage);
                return false;
            }

            values[fieldName] = parsedValue;
            return true;
        }

        /// <summary>
        /// 東日本シートのバリデーション結果をResultオブジェクトに反映
        /// </summary>
        private void ApplyEastValidationResult(ValidationResult validationResult, EastSheetValidationResult result)
        {
            // エラーメッセージの設定
            SetErrorIfExists(validationResult.Errors, "延実働車輌数", val => result.JitsudoError = val);
            SetErrorIfExists(validationResult.Errors, "搬送回数", val => result.HansoError = val);
            SetErrorIfExists(validationResult.Errors, "有料キロ数", val => result.YuryoKmError = val);
            SetErrorIfExists(validationResult.Errors, "無料キロ数", val => result.MuryoKmError = val);
            SetErrorIfExists(validationResult.Errors, "運輸実績", val => result.UnsoError = val);

            // 警告メッセージの設定（エラーがない場合のみ）
            SetWarningIfExists(validationResult.Warnings, "延実働車輌数", () => result.JitsudoError, val => result.JitsudoError = val);
            SetWarningIfExists(validationResult.Warnings, "搬送回数", () => result.HansoError, val => result.HansoError = val);
            SetWarningIfExists(validationResult.Warnings, "有料キロ数", () => result.YuryoKmError, val => result.YuryoKmError = val);
            SetWarningIfExists(validationResult.Warnings, "無料キロ数", () => result.MuryoKmError, val => result.MuryoKmError = val);
            SetWarningIfExists(validationResult.Warnings, "運輸実績", () => result.UnsoError, val => result.UnsoError = val);
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// 文字列を数値にパースする
        /// </summary>
        private bool TryParseValue(string input, out double? result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(input))
                return true;

            if (double.TryParse(input, out double parsedValue))
            {
                result = parsedValue;
                return true;
            }

            return false;
        }

        /// <summary>
        /// バリデーション結果のエラーを設定するヘルパーメソッド
        /// </summary>
        private void SetErrorIfExists(Dictionary<string, List<string>> errors, string key, Action<string> setError)
        {
            if (errors.ContainsKey(key))
            {
                setError(string.Join(", ", errors[key]));
            }
        }

        /// <summary>
        /// バリデーション結果の警告を設定するヘルパーメソッド（既存エラーがない場合のみ）
        /// </summary>
        private void SetWarningIfExists(
            Dictionary<string, List<string>> warnings,
            string key,
            Func<string> getCurrentError,
            Action<string> setWarning)
        {
            if (warnings.ContainsKey(key) && string.IsNullOrEmpty(getCurrentError()))
            {
                setWarning(WarningPrefix + string.Join(", ", warnings[key]));
            }
        }

        #endregion
    }

    #region Validation Result Classes

    /// <summary>
    /// 通常シートのバリデーション結果
    /// </summary>
    public class NormalSheetValidationResult
    {
        public string DayError { get; set; } = string.Empty;
        public string YuryoKmError { get; set; } = string.Empty;
        public string MuryoKmError { get; set; } = string.Empty;
        public string LateValueError { get; set; } = string.Empty;
        public bool HasErrors { get; set; }
    }

    /// <summary>
    /// 東日本シートのバリデーション結果
    /// </summary>
    public class EastSheetValidationResult
    {
        public string JitsudoError { get; set; } = string.Empty;
        public string HansoError { get; set; } = string.Empty;
        public string YuryoKmError { get; set; } = string.Empty;
        public string MuryoKmError { get; set; } = string.Empty;
        public string UnsoError { get; set; } = string.Empty;
        public bool HasErrors { get; set; }
    }

    #endregion
}