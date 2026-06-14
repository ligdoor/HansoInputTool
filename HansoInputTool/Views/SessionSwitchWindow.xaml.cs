using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class SessionSwitchWindow : Window
    {
        public SessionSwitchWindow(SessionSwitchViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
            vm.CloseDialog = result => DialogResult = result;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
