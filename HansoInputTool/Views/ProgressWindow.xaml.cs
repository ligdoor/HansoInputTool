using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class ProgressWindow : Window
    {
        public ProgressWindow(ProgressWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}