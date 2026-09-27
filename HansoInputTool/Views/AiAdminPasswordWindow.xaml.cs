using System.Windows;

namespace HansoInputTool.Views
{
    public partial class AiAdminPasswordWindow : Window
    {
        private readonly bool _setupMode;
        public string Password { get; private set; }

        public AiAdminPasswordWindow(bool setupMode)
        {
            InitializeComponent();
            _setupMode = setupMode;
            Title = setupMode ? "AI設定管理パスワードの初回設定" : "AI設定の認証";
            if (!setupMode)
            {
                ConfirmLabel.Visibility = Visibility.Collapsed;
                ConfirmPasswordBox.Visibility = Visibility.Collapsed;
                Height = 245;
            }
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e) { }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            var password = PasswordBox.Password;
            if (string.IsNullOrWhiteSpace(password) || password.Length < 6)
            {
                MessageBox.Show("AI設定管理パスワードは6文字以上で設定してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_setupMode && password != ConfirmPasswordBox.Password)
            {
                MessageBox.Show("確認用パスワードが一致しません。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Password = password;
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
