using Microsoft.Xaml.Behaviors;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


// WPF型を明示（WindowsAPICodePack経由のSystem.Windows.Forms競合を解消）
using Control      = System.Windows.Controls.Control;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using TextBox      = System.Windows.Controls.TextBox;
using DataObject   = System.Windows.DataObject;
namespace HansoInputTool.Behaviors
{
    public class EnterKeyTraversalBehavior : Behavior<Control>
    {
        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject != null)
            {
                AssociatedObject.KeyDown += OnKeyDown;
            }
        }

        protected override void OnDetaching()
        {
            if (AssociatedObject != null)
            {
                AssociatedObject.KeyDown -= OnKeyDown;
            }
            base.OnDetaching();
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && sender is UIElement element)
            {
                var request = new TraversalRequest(FocusNavigationDirection.Next);
                element.MoveFocus(request);
                e.Handled = true;
            }
        }
    }

    public class EnterKeyAndRegisterBehavior : Behavior<Control>
    {
        public static readonly DependencyProperty CommandProperty =
            DependencyProperty.Register(nameof(Command), typeof(ICommand), typeof(EnterKeyAndRegisterBehavior), new PropertyMetadata(null));

        public ICommand Command
        {
            get { return (ICommand)GetValue(CommandProperty); }
            set { SetValue(CommandProperty, value); }
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject != null)
            {
                AssociatedObject.KeyDown += OnKeyDown;
            }
        }

        protected override void OnDetaching()
        {
            if (AssociatedObject != null)
            {
                AssociatedObject.KeyDown -= OnKeyDown;
            }
            base.OnDetaching();
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Command != null && Command.CanExecute(null))
            {
                Command.Execute(null);
                e.Handled = true;
            }
        }
    }

    public class NumericInputBehavior : Behavior<TextBox>
    {
        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject != null)
            {
                // 数値入力欄ではIME（日本語入力）を使わない。変換中の状態や全角数字の混入を、そもそも起こさない。
                InputMethod.SetIsInputMethodEnabled(AssociatedObject, false);

                AssociatedObject.PreviewTextInput += OnPreviewTextInput;
                DataObject.AddPastingHandler(AssociatedObject, OnPasting);
            }
        }

        protected override void OnDetaching()
        {
            if (AssociatedObject != null)
            {
                AssociatedObject.PreviewTextInput -= OnPreviewTextInput;
                DataObject.RemovePastingHandler(AssociatedObject, OnPasting);
                AssociatedObject.ClearValue(InputMethod.IsInputMethodEnabledProperty);
            }
            base.OnDetaching();
        }

        private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 空の文字列で呼ばれることがある（その場合は何もしない）
            if (string.IsNullOrEmpty(e.Text)) return;

            // 半角の数字（0〜9）だけを受け付ける。全角数字や文字が混ざっていたら、入力全体を弾く
            if (!IsAsciiDigits(e.Text))
            {
                e.Handled = true;
            }
        }

        private void OnPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        private static bool IsAsciiDigits(string text)
        {
            foreach (char c in text)
                if (c < '0' || c > '9') return false;
            return true;
        }

        private static bool IsTextAllowed(string text)
        {
            return new System.Text.RegularExpressions.Regex("^[0-9]+$").IsMatch(text);
        }
    }
}