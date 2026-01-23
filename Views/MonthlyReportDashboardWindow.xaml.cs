using System.Windows;
using HansoInputTool.ViewModels;

namespace HansoInputTool.Views
{
    /// <summary>
    /// MonthlyReportDashboardWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MonthlyReportDashboardWindow : Window
    {
        public MonthlyReportDashboardWindow()
        {
            InitializeComponent();
            DataContext = new MonthlyReportDashboardViewModel();
        }
    }
}
