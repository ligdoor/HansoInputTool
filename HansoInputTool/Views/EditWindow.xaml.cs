using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Collections.Generic;
using System.Linq;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class EditWindow : Window
    {
        public EditWindow(EditWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            PreviewKeyDown += EditWindow_PreviewKeyDown;
        }

        /// <summary>
        /// 編集画面ではEnterで次の入力欄へ移動し、最後の入力欄でEnterを押すと保存する。
        /// 特殊フラグのCheckBoxや操作ボタンは通常のEnter移動対象から除外する。
        /// </summary>
        private void EditWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || Keyboard.FocusedElement is not TextBox current)
                return;

            // XAML上の視覚ツリー順＝業務上の入力順になるよう、TextBoxだけを拾う。
            // フラグのCheckBoxや削除・追加ボタンはTextBoxではないためEnter移動から除外される。
            var textBoxes = FindVisualChildren<TextBox>(this)
                .Where(tb => tb.IsVisible && tb.IsEnabled && tb.IsTabStop)
                .ToList();

            int currentIndex = textBoxes.IndexOf(current);
            if (currentIndex < 0) return;

            if (currentIndex < textBoxes.Count - 1)
            {
                textBoxes[currentIndex + 1].Focus();
                textBoxes[currentIndex + 1].SelectAll();
            }
            else if (DataContext is EditWindowViewModel vm && vm.SaveCommand.CanExecute(this))
            {
                vm.SaveCommand.Execute(this);
            }

            e.Handled = true;
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null) yield break;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T typed) yield return typed;
                foreach (var descendant in FindVisualChildren<T>(child))
                    yield return descendant;
            }
        }

    }
}