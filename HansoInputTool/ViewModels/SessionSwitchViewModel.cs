using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    public class SessionSwitchViewModel : ObservableObject
    {
        private readonly DatabaseService _dbService;

        public ObservableCollection<MonthSession> Sessions { get; } = new();

        private MonthSession _selectedSession;
        public MonthSession SelectedSession
        {
            get => _selectedSession;
            set
            {
                SetProperty(ref _selectedSession, value);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public ICommand SwitchCommand { get; }
        public ICommand DeleteCommand { get; }

        /// <summary>ダイアログを閉じるためのアクション（Viewからセット）</summary>
        public System.Action<bool?> CloseDialog { get; set; }

        /// <summary>切替後のセッションID（MainViewModelで参照）</summary>
        public long? SwitchedToSessionId { get; private set; }

        public SessionSwitchViewModel(DatabaseService dbService)
        {
            _dbService = dbService;

            SwitchCommand = new RelayCommand(
                _ => Switch(),
                _ => SelectedSession != null && SelectedSession.Id != _dbService.CurrentSessionId);

            DeleteCommand = new RelayCommand(
                _ => Delete(),
                _ => SelectedSession != null);

            Reload();
        }

        private void Reload()
        {
            Sessions.Clear();
            foreach (var s in _dbService.GetAllSessions())
                Sessions.Add(s);
        }

        private void Switch()
        {
            SwitchedToSessionId = SelectedSession.Id;
            CloseDialog?.Invoke(true);
        }

        private void Delete()
        {
            var result = System.Windows.MessageBox.Show(
                $"「{SelectedSession.Label}」のデータを削除します。\nこの操作は元に戻せません。よろしいですか？",
                "削除確認",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result != System.Windows.MessageBoxResult.Yes) return;

            _dbService.DeleteSession(SelectedSession.Id);
            Reload();
            CommandManager.InvalidateRequerySuggested();
        }
    }
}
