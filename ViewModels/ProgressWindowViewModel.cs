using HansoInputTool.ViewModels.Base;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace HansoInputTool.ViewModels
{
    public class ProgressWindowViewModel : ObservableObject
    {
        private string _statusText = "転記処理を開始しています...";
        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        private int _progressValue;
        public int ProgressValue
        {
            get => _progressValue;
            set => SetProperty(ref _progressValue, value);
        }

        private string _progressPercentage = "0%";
        public string ProgressPercentage
        {
            get => _progressPercentage;
            set => SetProperty(ref _progressPercentage, value);
        }

        private readonly StringBuilder _logBuilder = new();
        public string LogText => _logBuilder.ToString();

        private bool _isCloseButtonEnabled;
        public bool IsCloseButtonEnabled
        {
            get => _isCloseButtonEnabled;
            set => SetProperty(ref _isCloseButtonEnabled, value);
        }

        public ICommand CloseCommand { get; }

        public ProgressWindowViewModel()
        {
            CloseCommand = new RelayCommand(p => ((Window)p).Close());
        }

        public void UpdateProgress(int current, int total, string message)
        {
            if (total == 0) return;
            int percent = (int)(((double)current / total) * 100);
            ProgressValue = percent;
            ProgressPercentage = $"{percent}%";
            AppendLog(message);
        }

        public void AppendLog(string message)
        {
            _logBuilder.AppendLine(message);
            OnPropertyChanged(nameof(LogText));
        }

        public void Complete(string message = "全ての処理が完了しました。")
        {
            ProgressValue = 100;
            ProgressPercentage = "100%";
            StatusText = "✅ 処理が完了しました！";
            AppendLog(message);
            IsCloseButtonEnabled = true;
        }

        public void ErrorComplete(string message = "エラーにより処理を中断しました。")
        {
            StatusText = "❌ エラーが発生しました";
            AppendLog(message);
            IsCloseButtonEnabled = true;
        }
    }
}