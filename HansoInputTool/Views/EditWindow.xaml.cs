using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class EditWindow : Window
    {
        public EditWindow(EditWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}