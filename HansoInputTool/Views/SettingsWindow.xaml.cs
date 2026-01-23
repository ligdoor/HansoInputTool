using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow(SettingsWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}