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

        /// <summary>切替候補として表示するセッション一覧（クリア済みで0件のものは自動的に除外される）</summary>
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
        public ICommand ConfirmCommand { get; }
        public ICommand UnconfirmCommand { get; }

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

            ConfirmCommand = new RelayCommand(
                _ => Confirm(),
                _ => SelectedSession != null && !SelectedSession.IsConfirmed);

            UnconfirmCommand = new RelayCommand(
                _ => Unconfirm(),
                _ => SelectedSession != null && SelectedSession.IsConfirmed);

            Reload();
        }

        /// <summary>
        /// 選択中のセッションを「確定」にする。確定後は誤操作での編集・削除を防ぐため
        /// このアプリからの入力・更新・削除・クリアがブロックされる（確定解除するまで）。
        /// </summary>
        private void Confirm()
        {
            var result = System.Windows.MessageBox.Show(
                $"「{SelectedSession.Label}」を確定します。\n確定後は解除するまで編集・削除できなくなります。よろしいですか？",
                "確定確認",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Question);

            if (result != System.Windows.MessageBoxResult.Yes) return;

            _dbService.ConfirmSession(SelectedSession.Id);
            Reload();
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>選択中のセッションの確定を解除し、再び編集できるようにする</summary>
        private void Unconfirm()
        {
            var result = System.Windows.MessageBox.Show(
                $"「{SelectedSession.Label}」の確定を解除します。\n再び編集できる状態に戻ります。よろしいですか？",
                "確定解除の確認",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result != System.Windows.MessageBoxResult.Yes) return;

            _dbService.UnconfirmSession(SelectedSession.Id);
            Reload();
            CommandManager.InvalidateRequerySuggested();
        }

        /// <summary>
        /// セッション一覧を再取得する。その前に、今使っている月以外で搬送データが
        /// 0件になっているセッション（クリア済みで不要になった月）を自動的に削除する。
        /// </summary>
        private void Reload()
        {
            _dbService.CleanUpEmptySessions();
            Sessions.Clear();
            foreach (var s in _dbService.GetAllSessions())
                Sessions.Add(s);
        }

        private void Switch()
        {
            SwitchedToSessionId = SelectedSession.Id;
            CloseDialog?.Invoke(true);
        }

        /// <summary>
        /// 選択中のセッションを削除する。
        /// もし削除対象が現在アクティブなセッションだった場合、DatabaseService側で
        /// 自動的に別のセッションへ切り替わる（DeleteSession内部の処理）。この場合、
        /// SwitchedToSessionIdはセットされない（明示的な「切替」操作ではないため）が、
        /// 呼び出し元のMainViewModel.OpenSessionSwitch()側でダイアログが閉じた後に
        /// CurrentSessionIdの変化を検知し、画面表示を追従させる。
        /// </summary>
        private void Delete()
        {
            if (SelectedSession.IsConfirmed)
            {
                System.Windows.MessageBox.Show(
                    $"「{SelectedSession.Label}」は確定済みのため削除できません。\n削除するには先に「確定解除」してください。",
                    "削除できません",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

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
