using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HansoInputTool.Messaging;
using HansoInputTool.ViewModels;


// WPF型を明示（WindowsAPICodePack経由のSystem.Windows.Forms競合を解消）
using Control      = System.Windows.Controls.Control;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using TextBox      = System.Windows.Controls.TextBox;
using DataObject   = System.Windows.DataObject;
namespace HansoInputTool.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var viewModel = new MainViewModel();

            // メッセンジャーを購読して、FocusMessageを受け取った時の動作を定義
            Messenger.Register<FocusMessage>(this, message =>
            {
                if (FindName(message.TargetElementName) is UIElement targetElement)
                {
                    targetElement.Focus();
                }
            });

            // ショートカットキー処理
            this.PreviewKeyDown += MainWindow_PreviewKeyDown;

            DataContext = viewModel;
        }

        /// <summary>
        /// ショートカットキーの処理
        /// </summary>
        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // テキストボックスにフォーカスがある場合は、通常の入力を優先
            if (Keyboard.FocusedElement is TextBox textBox)
            {
                // 修飾キーがある場合のみショートカットとして処理
                if (Keyboard.Modifiers == ModifierKeys.None)
                    return;
            }
            
            if (DataContext is MainViewModel vm)
            {
                var key = e.Key == Key.System ? e.SystemKey : e.Key;
                var modifiers = Keyboard.Modifiers;
                
                if (vm.ProcessShortcut(key, modifiers))
                {
                    e.Handled = true;
                }
            }
        }

        // Enterキーで次のコントロールに移動する処理
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var request = new TraversalRequest(FocusNavigationDirection.Next);
                if (Keyboard.FocusedElement is UIElement elementWithFocus)
                {
                    elementWithFocus.MoveFocus(request);
                }
                e.Handled = true;
            }
        }

        // 通常シートの最後の入力欄でEnterキーを押したら登録する処理
        private void LastNormalTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (DataContext is MainViewModel vm && vm.RegisterNormalCommand.CanExecute(null))
                {
                    vm.RegisterNormalCommand.Execute(null);
                }
                e.Handled = true;
            }
        }

        // 東日本シートの最後の入力欄でEnterキーを押したら登録する処理
        private void LastEastTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (DataContext is MainViewModel vm && vm.RegisterEastCommand.CanExecute(null))
                {
                    vm.RegisterEastCommand.Execute(null);
                }
                e.Handled = true;
            }
        }

        /// <summary>ログ末尾に自動スクロール</summary>
        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb)
                tb.ScrollToEnd();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}
