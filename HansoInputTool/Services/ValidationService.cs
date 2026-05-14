using System;
using System.Collections.Generic;
using System.Linq;

namespace HansoInputTool.Services
{
    public class ValidationService
    {
        /// <summary>
        /// 通常シートのデータを検証
        /// </summary>
        /// <param name="values">入力値</param>
        /// <param name="sheetName">シート名</param>
        /// <param name="year">入力対象の年（月末日チェックに使用。0なら現在年で判定）</param>
        /// <param name="month">入力対象の月（月末日チェックに使用。0なら1〜31の範囲チェックのみ）</param>
        public ValidationResult ValidateNormalData(
            Dictionary<string, double?> values,
            string sheetName,
            int year  = 0,
            int month = 0)
        {
            var result = new ValidationResult();

            // 日付の検証
            if (!values.ContainsKey("日(B)") || !values["日(B)"].HasValue)
            {
                result.AddError("日付", "日付は必須です");
            }
            else
            {
                var day = (int)values["日(B)"].Value;

                if (day < 1 || day > 31)
                {
                    result.AddError("日付", "日付は1-31の範囲で入力してください");
                }
                else if (month >= 1 && month <= 12)
                {
                    // 月が分かっている場合は実際の月末日と照合する
                    int useYear = (year >= 1900) ? year : DateTime.Now.Year;
                    int lastDay = DateTime.DaysInMonth(useYear, month);
                    if (day > lastDay)
                    {
                        result.AddError("日付",
                            $"{month}月は{lastDay}日までです（入力値: {day}日）");
                    }
                }
            }

            // 有料キロの検証
            if (values.ContainsKey("有料キロ(D)") && values["有料キロ(D)"].HasValue)
            {
                var km = values["有料キロ(D)"].Value;
                if (km < 0)
                {
                    result.AddError("有料キロ", "有料キロは0以上で入力してください");
                }
                else if (km > 500)
                {
                    result.AddWarning("有料キロ", "有料キロが500kmを超えています。正しいですか？");
                }
                else if (km > 0 && km < 1)
                {
                    result.AddWarning("有料キロ", "有料キロが1km未満です。正しいですか？");
                }
            }

            // 無料キロの検証
            if (values.ContainsKey("無料キロ(E)") && values["無料キロ(E)"].HasValue)
            {
                var km = values["無料キロ(E)"].Value;
                if (km < 0)
                {
                    result.AddError("無料キロ", "無料キロは0以上で入力してください");
                }
                else if (km > 200)
                {
                    result.AddWarning("無料キロ", "無料キロが200kmを超えています。正しいですか？");
                }
            }

            // 深夜時間の検証（大月以外）
            if (!sheetName.Contains("大月") && values.ContainsKey("深夜時間(K)") && values["深夜時間(K)"].HasValue)
            {
                var minutes = values["深夜時間(K)"].Value;
                if (minutes < 0)
                {
                    result.AddError("深夜時間", "深夜時間は0以上で入力してください");
                }
                else if (minutes > 1440)
                {
                    result.AddError("深夜時間", "深夜時間は1440分(24時間)を超えることはできません");
                }
                else if (minutes > 720)
                {
                    result.AddWarning("深夜時間", "深夜時間が12時間を超えています。正しいですか？");
                }
            }

            // 深夜料金の検証（大月のみ）
            if (sheetName.Contains("大月") && values.ContainsKey("深夜料金(H)") && values["深夜料金(H)"].HasValue)
            {
                var fee = values["深夜料金(H)"].Value;
                if (fee < 0)
                {
                    result.AddError("深夜料金", "深夜料金は0以上で入力してください");
                }
                else if (fee > 50000)
                {
                    result.AddWarning("深夜料金", "深夜料金が50,000円を超えています。正しいですか？");
                }
            }

            // 有料キロと無料キロの合計チェック
            var totalKm = (values.GetValueOrDefault("有料キロ(D)") ?? 0) +
                         (values.GetValueOrDefault("無料キロ(E)") ?? 0);
            if (totalKm > 1000)
            {
                result.AddWarning("走行距離", $"合計走行距離が{totalKm:F1}kmです。正しいですか？");
            }

            return result;
        }

        /// <summary>
        /// 東日本シートのデータを検証
        /// </summary>
        public ValidationResult ValidateEastData(Dictionary<string, double?> values)
        {
            var result = new ValidationResult();

            // 延実働車輌数の検証
            if (values.ContainsKey("延実働車輌数") && values["延実働車輌数"].HasValue)
            {
                var count = values["延実働車輌数"].Value;
                if (count < 0)
                {
                    result.AddError("延実働車輌数", "延実働車輌数は0以上で入力してください");
                }
                else if (count > 100)
                {
                    result.AddWarning("延実働車輌数", "延実働車輌数が100を超えています。正しいですか？");
                }
            }

            // 搬送回数の検証
            if (values.ContainsKey("搬送回数") && values["搬送回数"].HasValue)
            {
                var count = values["搬送回数"].Value;
                if (count < 0)
                {
                    result.AddError("搬送回数", "搬送回数は0以上で入力してください");
                }
                else if (count > 500)
                {
                    result.AddWarning("搬送回数", "搬送回数が500を超えています。正しいですか？");
                }
            }

            // 有料キロ数の検証
            if (values.ContainsKey("有料キロ数") && values["有料キロ数"].HasValue)
            {
                var km = values["有料キロ数"].Value;
                if (km < 0)
                {
                    result.AddError("有料キロ数", "有料キロ数は0以上で入力してください");
                }
                else if (km > 10000)
                {
                    result.AddWarning("有料キロ数", "有料キロ数が10,000kmを超えています。正しいですか？");
                }
            }

            // 無料キロ数の検証
            if (values.ContainsKey("無料キロ数") && values["無料キロ数"].HasValue)
            {
                var km = values["無料キロ数"].Value;
                if (km < 0)
                {
                    result.AddError("無料キロ数", "無料キロ数は0以上で入力してください");
                }
                else if (km > 5000)
                {
                    result.AddWarning("無料キロ数", "無料キロ数が5,000kmを超えています。正しいですか？");
                }
            }

            // 運輸実績の検証
            if (values.ContainsKey("運輸実績") && values["運輸実績"].HasValue)
            {
                var amount = values["運輸実績"].Value;
                if (amount < 0)
                {
                    result.AddError("運輸実績", "運輸実績は0以上で入力してください");
                }
                else if (amount > 10000000)
                {
                    result.AddWarning("運輸実績", "運輸実績が10,000,000円を超えています。正しいですか？");
                }
            }

            return result;
        }
    }

    public class ValidationResult
    {
        public Dictionary<string, List<string>> Errors { get; } = new Dictionary<string, List<string>>();
        public Dictionary<string, List<string>> Warnings { get; } = new Dictionary<string, List<string>>();

        public bool IsValid => !Errors.Any();
        public bool HasWarnings => Warnings.Any();

        public void AddError(string field, string message)
        {
            if (!Errors.ContainsKey(field))
                Errors[field] = new List<string>();
            Errors[field].Add(message);
        }

        public void AddWarning(string field, string message)
        {
            if (!Warnings.ContainsKey(field))
                Warnings[field] = new List<string>();
            Warnings[field].Add(message);
        }

        public string GetErrorMessage()
        {
            return string.Join("\n", Errors.SelectMany(e =>
                e.Value.Select(msg => $"【{e.Key}】{msg}")));
        }

        public string GetWarningMessage()
        {
            return string.Join("\n", Warnings.SelectMany(w =>
                w.Value.Select(msg => $"【{w.Key}】{msg}")));
        }
    }
}