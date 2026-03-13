using System.Windows;

namespace HansoInputTool.Views
{
    /// <summary>
    /// Claude APIキーを初回入力するシンプルなダイアログ
    /// </summary>
    public partial class ApiKeyInputWindow : Window
    {
        public string ApiKey { get; private set; }

        public ApiKeyInputWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ApiKey = ApiKeyBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(ApiKey))
            {
                MessageBox.Show("APIキーを入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
