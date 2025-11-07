using Microsoft.Xaml.Behaviors;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
            }
            base.OnDetaching();
        }

        private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!char.IsDigit(e.Text, e.Text.Length - 1))
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

        private static bool IsTextAllowed(string text)
        {
            return new System.Text.RegularExpressions.Regex("^[0-9]+$").IsMatch(text);
        }
    }
}