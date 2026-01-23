// Views/TransferConfirmationWindow.xaml.cs
using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class TransferConfirmationWindow : Window
    {
        public TransferConfirmationWindow(TransferConfirmationViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}