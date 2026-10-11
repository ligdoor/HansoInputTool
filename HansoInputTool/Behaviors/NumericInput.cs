using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

// WPF型を明示（WindowsAPICodePack経由のSystem.Windows.Forms競合を解消）
using DataObject = System.Windows.DataObject;
using TextBox    = System.Windows.Controls.TextBox;

namespace HansoInputTool.Behaviors
{
    /// <summary>数値入力欄の入力制限の種類</summary>
    public enum NumericInputMode
    {
        /// <summary>制限なし（既定）</summary>
        None,

        /// <summary>半角の数字（0〜9）だけ。IMEは使わない</summary>
        Integer,

        /// <summary>半角の数字と小数点（.）1つまで。IMEは使わない</summary>
        Decimal,

        /// <summary>文字の制限はしないが、IME（日本語入力）は使わない（列記号など半角英数字の欄用）</summary>
        ImeOff,
    }

    /// <summary>
    /// TextBox に <c>NumericInput.Mode="Integer"</c> のように付けるだけで、
    /// 数値入力欄として動くようにする（IMEを自動でオフにし、半角の数字以外は入力・貼り付けできなくする）。
    /// </summary>
    public static class NumericInput
    {
        public static readonly DependencyProperty ModeProperty =
            DependencyProperty.RegisterAttached(
                "Mode", typeof(NumericInputMode), typeof(NumericInput),
                new PropertyMetadata(NumericInputMode.None, OnModeChanged));

        public static NumericInputMode GetMode(DependencyObject d) => (NumericInputMode)d.GetValue(ModeProperty);
        public static void SetMode(DependencyObject d, NumericInputMode value) => d.SetValue(ModeProperty, value);

        private static readonly Regex IntegerPattern = new(@"^[0-9]*\z", RegexOptions.Compiled);
        private static readonly Regex DecimalPattern = new(@"^[0-9]*\.?[0-9]*\z", RegexOptions.Compiled);

        private static void OnModeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBox tb) return;

            // 付け直しに備えて、いったん外す
            tb.PreviewTextInput -= OnPreviewTextInput;
            DataObject.RemovePastingHandler(tb, OnPasting);

            var mode = (NumericInputMode)e.NewValue;
            if (mode == NumericInputMode.None)
            {
                tb.ClearValue(InputMethod.IsInputMethodEnabledProperty);
                return;
            }

            // 日本語入力をオフ（変換中の状態や全角文字の混入を、そもそも起こさない）
            InputMethod.SetIsInputMethodEnabled(tb, false);

            if (mode == NumericInputMode.Integer || mode == NumericInputMode.Decimal)
            {
                tb.PreviewTextInput += OnPreviewTextInput;
                DataObject.AddPastingHandler(tb, OnPasting);
            }
        }

        private static void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 空の文字列で呼ばれることがある（その場合は何もしない）
            if (string.IsNullOrEmpty(e.Text)) return;

            e.Handled = !IsAllowed((TextBox)sender, e.Text);
        }

        private static void OnPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            var text = (string)e.DataObject.GetData(typeof(string));
            if (string.IsNullOrEmpty(text) || !IsAllowed((TextBox)sender, text))
                e.CancelCommand();
        }

        /// <summary>入力（または貼り付け）を反映した後の文字列が、許可された形かどうか</summary>
        private static bool IsAllowed(TextBox tb, string input)
        {
            string current = tb.Text ?? string.Empty;
            int start  = Math.Min(tb.SelectionStart, current.Length);
            int length = Math.Min(tb.SelectionLength, current.Length - start);
            string next = current.Remove(start, length).Insert(start, input);

            return GetMode(tb) == NumericInputMode.Integer
                ? IntegerPattern.IsMatch(next)
                : DecimalPattern.IsMatch(next);
        }
    }
}
