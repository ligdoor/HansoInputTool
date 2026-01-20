using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HansoInputTool.Controls
{
    /// <summary>
    /// ショートカットキーを入力するためのカスタムテキストボックス
    /// </summary>
    public class ShortcutKeyBox : TextBox
    {
        public static readonly DependencyProperty ShortcutKeyProperty =
            DependencyProperty.Register(
                nameof(ShortcutKey),
                typeof(Key),
                typeof(ShortcutKeyBox),
                new FrameworkPropertyMetadata(Key.None, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnShortcutChanged));

        public static readonly DependencyProperty ShortcutModifiersProperty =
            DependencyProperty.Register(
                nameof(ShortcutModifiers),
                typeof(ModifierKeys),
                typeof(ShortcutKeyBox),
                new FrameworkPropertyMetadata(ModifierKeys.None, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnShortcutChanged));

        public static readonly DependencyProperty IsRecordingProperty =
            DependencyProperty.Register(
                nameof(IsRecording),
                typeof(bool),
                typeof(ShortcutKeyBox),
                new PropertyMetadata(false, OnIsRecordingChanged));

        public Key ShortcutKey
        {
            get => (Key)GetValue(ShortcutKeyProperty);
            set => SetValue(ShortcutKeyProperty, value);
        }

        public ModifierKeys ShortcutModifiers
        {
            get => (ModifierKeys)GetValue(ShortcutModifiersProperty);
            set => SetValue(ShortcutModifiersProperty, value);
        }

        public bool IsRecording
        {
            get => (bool)GetValue(IsRecordingProperty);
            set => SetValue(IsRecordingProperty, value);
        }

        private static readonly SolidColorBrush RecordingBrush = new SolidColorBrush(Color.FromRgb(59, 130, 246));
        private static readonly SolidColorBrush NormalBrush = new SolidColorBrush(Color.FromRgb(23, 32, 51));

        public ShortcutKeyBox()
        {
            IsReadOnly = true;
            Cursor = Cursors.Hand;
            TextAlignment = TextAlignment.Center;
            UpdateDisplayText();
        }

        private static void OnShortcutChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ShortcutKeyBox box)
            {
                box.UpdateDisplayText();
            }
        }

        private static void OnIsRecordingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ShortcutKeyBox box)
            {
                box.BorderBrush = box.IsRecording ? RecordingBrush : NormalBrush;
                box.BorderThickness = box.IsRecording ? new Thickness(2) : new Thickness(1);
                
                if (box.IsRecording)
                {
                    box.Text = "キーを押してください...";
                }
                else
                {
                    box.UpdateDisplayText();
                }
            }
        }

        protected override void OnGotFocus(RoutedEventArgs e)
        {
            base.OnGotFocus(e);
            IsRecording = true;
        }

        protected override void OnLostFocus(RoutedEventArgs e)
        {
            base.OnLostFocus(e);
            IsRecording = false;
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (!IsRecording)
            {
                base.OnPreviewKeyDown(e);
                return;
            }

            e.Handled = true;

            // Escapeで録音をキャンセル
            if (e.Key == Key.Escape)
            {
                IsRecording = false;
                Keyboard.ClearFocus();
                return;
            }

            // 修飾キーのみの場合は無視
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl ||
                e.Key == Key.LeftAlt || e.Key == Key.RightAlt ||
                e.Key == Key.LeftShift || e.Key == Key.RightShift ||
                e.Key == Key.LWin || e.Key == Key.RWin ||
                e.Key == Key.System)
            {
                return;
            }

            // 実際のキーを取得（System修飾子の場合はSystemKeyを使用）
            var key = e.Key == Key.System ? e.SystemKey : e.Key;

            // 修飾キーを取得
            var modifiers = Keyboard.Modifiers;

            // 修飾キーなしの場合、一部のキーは許可しない（誤操作防止）
            if (modifiers == ModifierKeys.None)
            {
                // アルファベット、数字キー単独は許可しない
                if ((key >= Key.A && key <= Key.Z) || (key >= Key.D0 && key <= Key.D9))
                {
                    return;
                }
            }

            ShortcutKey = key;
            ShortcutModifiers = modifiers;
            IsRecording = false;
            Keyboard.ClearFocus();
        }

        private void UpdateDisplayText()
        {
            if (ShortcutKey == Key.None)
            {
                Text = "(クリックして設定)";
                return;
            }

            var parts = new System.Collections.Generic.List<string>();

            if (ShortcutModifiers.HasFlag(ModifierKeys.Control))
                parts.Add("Ctrl");
            if (ShortcutModifiers.HasFlag(ModifierKeys.Alt))
                parts.Add("Alt");
            if (ShortcutModifiers.HasFlag(ModifierKeys.Shift))
                parts.Add("Shift");

            parts.Add(GetKeyDisplayName(ShortcutKey));

            Text = string.Join("+", parts);
        }

        private static string GetKeyDisplayName(Key key)
        {
            return key switch
            {
                Key.OemComma => ",",
                Key.OemPeriod => ".",
                Key.OemPlus => "+",
                Key.OemMinus => "-",
                Key.OemQuestion => "/",
                Key.OemOpenBrackets => "[",
                Key.OemCloseBrackets => "]",
                Key.OemSemicolon => ";",
                Key.OemQuotes => "'",
                Key.Left => "←",
                Key.Right => "→",
                Key.Up => "↑",
                Key.Down => "↓",
                Key.Enter => "Enter",
                Key.Tab => "Tab",
                Key.Delete => "Delete",
                Key.Back => "Backspace",
                Key.Escape => "Esc",
                Key.Space => "Space",
                Key.F1 => "F1",
                Key.F2 => "F2",
                Key.F3 => "F3",
                Key.F4 => "F4",
                Key.F5 => "F5",
                Key.F6 => "F6",
                Key.F7 => "F7",
                Key.F8 => "F8",
                Key.F9 => "F9",
                Key.F10 => "F10",
                Key.F11 => "F11",
                Key.F12 => "F12",
                _ => key.ToString()
            };
        }
    }
}
