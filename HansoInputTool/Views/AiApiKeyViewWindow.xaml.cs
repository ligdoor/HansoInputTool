using System;
using System.Windows;
using System.Windows.Threading;

namespace HansoInputTool.Views
{
    public partial class AiApiKeyViewWindow : Window
    {
        private readonly DispatcherTimer _timer;
        private int _remainingSeconds = 30;

        public AiApiKeyViewWindow(string apiKey)
        {
            InitializeComponent();
            DataContext = new { ApiKey = apiKey };
            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _remainingSeconds--;
            CountdownText.Text = $"残り{_remainingSeconds}秒";
            if (_remainingSeconds <= 0)
            {
                _timer.Stop();
                Close();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        protected override void OnClosed(EventArgs e)
        {
            _timer.Stop();
            base.OnClosed(e);
        }
    }
}
