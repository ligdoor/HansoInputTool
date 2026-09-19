using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class LateNightCalculatorWindow : Window
    {
        public LateNightCalculatorWindow(LateNightCalculatorViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            Loaded += (_, _) => CompanyHourTextBox.Focus();
            PreviewKeyDown += LateNightCalculatorWindow_PreviewKeyDown;
        }

        /// <summary>
        /// Enterキーで次の入力欄へ移動する。最後の入力欄（故人宅出発・分）でEnterを押した場合は
        /// 「計算する」を実行する（EditWindowのEnter移動と同じ考え方）。
        /// </summary>
        private void LateNightCalculatorWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || Keyboard.FocusedElement is not TextBox current)
                return;

            var textBoxes = new[] { CompanyHourTextBox, CompanyMinuteTextBox, HomeHourTextBox, HomeMinuteTextBox };
            int currentIndex = System.Array.IndexOf(textBoxes, current);
            if (currentIndex < 0) return;

            if (currentIndex < textBoxes.Length - 1)
            {
                textBoxes[currentIndex + 1].Focus();
                textBoxes[currentIndex + 1].SelectAll();
            }
            else if (DataContext is LateNightCalculatorViewModel vm && vm.CalculateCommand.CanExecute(null))
            {
                vm.CalculateCommand.Execute(null);
            }

            e.Handled = true;
        }
    }
}
