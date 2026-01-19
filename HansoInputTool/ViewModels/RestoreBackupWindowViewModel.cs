using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class RestoreBackupWindowViewModel : ObservableObject
    {
        private readonly BackupService _backupService;
        private readonly string _inputFilePath;
        private readonly string _templateFilePath;
        private readonly MainViewModel _mainViewModel;

        public ObservableCollection<BackupInfo> InputBackups { get; }
        public ObservableCollection<BackupInfo> TemplateBackups { get; }

        private BackupInfo _selectedInputBackup;
        public BackupInfo SelectedInputBackup
        {
            get => _selectedInputBackup;
            set => SetProperty(ref _selectedInputBackup, value);
        }

        private BackupInfo _selectedTemplateBackup;
        public BackupInfo SelectedTemplateBackup
        {
            get => _selectedTemplateBackup;
            set => SetProperty(ref _selectedTemplateBackup, value);
        }

        public ICommand RestoreCommand { get; }
        public ICommand OpenBackupFolderCommand { get; }
        public ICommand CancelCommand { get; }

        public RestoreBackupWindowViewModel(
            BackupService backupService,
            string inputFilePath,
            string templateFilePath,
            MainViewModel mainViewModel)
        {
            _backupService = backupService;
            _inputFilePath = inputFilePath;
            _templateFilePath = templateFilePath;
            _mainViewModel = mainViewModel;

            InputBackups = new ObservableCollection<BackupInfo>(
                _backupService.GetAvailableBackups("Input"));
            TemplateBackups = new ObservableCollection<BackupInfo>(
                _backupService.GetAvailableBackups("Template"));

            RestoreCommand = new RelayCommand(p => RestoreBackup(p));
            OpenBackupFolderCommand = new RelayCommand(p => _backupService.OpenBackupFolder());
            CancelCommand = new RelayCommand(p => ((Window)p).Close());
        }

        private void RestoreBackup(object parameter)
        {
            if (SelectedInputBackup == null && SelectedTemplateBackup == null)
            {
                MessageBox.Show("復元するバックアップを選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                "選択したバックアップから復元しますか？\n現在のファイルは上書きされます。",
                "復元確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            bool success = true;

            if (SelectedInputBackup != null)
            {
                success &= _backupService.RestoreFromBackup(SelectedInputBackup.FilePath, _inputFilePath);
            }

            if (SelectedTemplateBackup != null)
            {
                success &= _backupService.RestoreFromBackup(SelectedTemplateBackup.FilePath, _templateFilePath);
            }

            if (success)
            {
                MessageBox.Show("バックアップから復元しました。", "復元完了", MessageBoxButton.OK, MessageBoxImage.Information);
                _mainViewModel.ReloadAfterRestore();
                ((Window)parameter).Close();
            }
            else
            {
                MessageBox.Show("復元に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}