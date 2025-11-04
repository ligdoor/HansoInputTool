using System.Windows;
using System.Windows.Input;
using HansoInputTool.Messaging; // Messengerを使うために追加
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // ViewModelのインスタンスを先に作成
            var viewModel = new MainViewModel();

            // メッセンジャーを購読して、FocusMessageを受け取った時の動作を定義
            Messenger.Register<FocusMessage>(this, message =>
            {
                // メッセージで指定された名前のコントロールを探してフォーカスを当てる
                if (FindName(message.TargetElementName) is UIElement targetElement)
                {
                    targetElement.Focus();
                }
            });

            // DataContextに設定
            DataContext = viewModel;
        }

        // Enterキーで次のコントロールに移動する処理（これは変更なし）
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

        // 通常シートの最後の入力欄でEnterキーを押したら登録する処理（これは変更なし）
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

        // 東日本シートの最後の入力欄でEnterキーを押したら登録する処理（これは変更なし）
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
    }
}