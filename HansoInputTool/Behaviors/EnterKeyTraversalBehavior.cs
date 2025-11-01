using Microsoft.Xaml.Behaviors;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HansoInputTool.Behaviors
{
    // Enterキーで次のコントロールにフォーカスを移すビヘイビア
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

    // 最後の入力欄でEnterキーを押したらコマンドを実行するビヘイビア
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
}