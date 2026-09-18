using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HansoInputTool.Messaging;
using HansoInputTool.ViewModels;

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


        /// <summary>実績一覧でEnterを押すと選択行を編集、Deleteで削除。</summary>
        private void PreviewDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (DataContext is not MainViewModel vm || vm.IsBusy || vm.SelectedRow == null)
                return;

            if (e.Key == Key.Enter)
            {
                if (vm.EditRowCommand.CanExecute(null))
                    vm.EditRowCommand.Execute(null);
                e.Handled = true;
            }
            else if (e.Key == Key.Delete)
            {
                if (vm.DeleteRowCommand.CanExecute(null))
                    vm.DeleteRowCommand.Execute(null);
                e.Handled = true;
            }
        }

        /// <summary>実績一覧をダブルクリックしたら、その行を編集する。</summary>
        private void PreviewDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid || DataContext is not MainViewModel vm || vm.IsBusy)
                return;

            if (e.OriginalSource is DependencyObject source)
            {
                var row = ItemsControl.ContainerFromElement(grid, source) as DataGridRow;
                if (row == null) return;
                grid.SelectedItem = row.Item;
                if (vm.EditRowCommand.CanExecute(null))
                    vm.EditRowCommand.Execute(null);
                e.Handled = true;
            }
        }

        /// <summary>その他メニューを左クリックで開く。</summary>
        private void OtherButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
