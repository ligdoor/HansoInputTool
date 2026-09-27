using System;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;
using HansoInputTool.Views;

namespace HansoInputTool.ViewModels
{
    public partial class SettingsWindowViewModel
    {
        private bool _isAiSettingsUnlocked;

        public bool IsAiSettingsUnlocked
        {
            get => _isAiSettingsUnlocked;
            private set
            {
                if (!SetProperty(ref _isAiSettingsUnlocked, value)) return;
                OnPropertyChanged(nameof(AiSecurityStatus));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string AiSecurityStatus => IsAiSettingsUnlocked ? "🔓 編集可能" : "🔒 保護中";

        public ICommand UnlockAiSettingsCommand { get; private set; }
        public ICommand ChangeAiApiKeyCommand { get; private set; }
        public ICommand ViewAiApiKeyCommand { get; private set; }

        private void InitializeAiSecurityCommands()
        {
            UnlockAiSettingsCommand = new RelayCommand(_ => UnlockAiSettings(), _ => true);
            ChangeAiApiKeyCommand = new RelayCommand(p => ChangeAiApiKey(p), _ => IsAiSettingsUnlocked);
            ViewAiApiKeyCommand = new RelayCommand(p => ViewAiApiKey(p), _ => IsAiSettingsUnlocked && !string.IsNullOrWhiteSpace(_aiSettings?.ApiKey));
        }

        private void UnlockAiSettings()
        {
            try
            {
                if (!_aiSettingsService.HasAdminPassword())
                {
                    var setup = new AiAdminPasswordWindow(true)
                    {
                        Owner = Application.Current?.MainWindow
                    };
                    if (setup.ShowDialog() != true) return;

                    _aiSettingsService.SetAdminPassword(setup.Password);
                    IsAiSettingsUnlocked = true;
                    AiStatus = "管理認証済み";
                    return;
                }

                var dialog = new AiAdminPasswordWindow(false)
                {
                    Owner = Application.Current?.MainWindow
                };
                if (dialog.ShowDialog() != true) return;

                if (!_aiSettingsService.VerifyAdminPassword(dialog.Password))
                {
                    MessageBox.Show("AI設定管理パスワードが正しくありません。", "認証失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                IsAiSettingsUnlocked = true;
                AiStatus = "管理認証済み";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"AI設定の認証処理に失敗しました。\n\n{ex.Message}", "認証エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ChangeAiApiKey(object parameter)
        {
            if (!IsAiSettingsUnlocked) return;

            var dialog = new ApiKeyInputWindow(AiProvider)
            {
                Owner = parameter as Window ?? Application.Current?.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            _aiSettings.ApiKey = dialog.ApiKey;
            AiApiKey = dialog.ApiKey;
            _aiSettingsDirty = true;
            AiStatus = "APIキー変更あり（未保存）";
            CommandManager.InvalidateRequerySuggested();
        }

        private void ViewAiApiKey(object parameter)
        {
            if (!IsAiSettingsUnlocked || string.IsNullOrWhiteSpace(_aiSettings?.ApiKey)) return;

            var auth = new AiAdminPasswordWindow(false)
            {
                Owner = parameter as Window ?? Application.Current?.MainWindow
            };
            if (auth.ShowDialog() != true || !_aiSettingsService.VerifyAdminPassword(auth.Password))
            {
                MessageBox.Show("AI設定管理パスワードが正しくありません。", "認証失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new AiApiKeyViewWindow(_aiSettings.ApiKey)
            {
                Owner = parameter as Window ?? Application.Current?.MainWindow
            };
            dialog.ShowDialog();
        }
    }
}
