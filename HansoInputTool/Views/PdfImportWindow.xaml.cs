using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    public partial class PdfImportWindow : Window
    {
        public PdfImportWindow(PdfImportViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
