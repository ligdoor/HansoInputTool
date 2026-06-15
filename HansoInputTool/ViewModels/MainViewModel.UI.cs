using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.Services;
using Newtonsoft.Json.Linq;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel
    {
        #region ショートカット

        public bool ProcessShortcut(Key key, ModifierKeys modifiers)
        {
            if (_shortcutService == null) return false;
            foreach (var kvp in _shortcutService.CurrentSettings.Shortcuts)
                if (kvp.Value.Matches(key, modifiers))
                    return ExecuteShortcutAction(kvp.Key);
            return false;
        }

        private bool ExecuteShortcutAction(string actionName)
        {
            if (IsBusy) return false;
            switch (actionName)
            {
                case "Save":         return TryExecute(SaveInputCommand);
                case "Register":     return SelectedTabIndex == 0 ? TryExecute(RegisterNormalCommand) : TryExecute(RegisterEastCommand);
                case "NextSheet":    MoveSheet(+1); return true;
                case "PrevSheet":    MoveSheet(-1); return true;
                case "Transfer":     return TryExecute(TransferCommand);
                case "OpenSettings": return TryExecute(OpenSettingsCommand);
                case "SwitchTab":    SelectedTabIndex = (SelectedTabIndex + 1) % 2; return true;
                case "EditRow":      return TryExecute(EditRowCommand);
                case "DeleteRow":    return TryExecute(DeleteRowCommand);
                case "CreateBackup": return TryExecute(CreateBackupCommand);
                default:
                    // Flag_{flagId} 形式のアクション → 対応フラグをトグル
                    if (actionName.StartsWith("Flag_"))
                    {
                        var flagId = actionName.Substring(5);
                        return NormalSheet.ToggleFlag(flagId);
                    }
                    break;
            }
            return false;
        }

        /// <summary>
        /// フラグ定義のショートカット設定を ShortcutService に同期する。
        /// フラグ追加・削除・変更時に呼ぶ。
        /// </summary>
        public void SyncFlagShortcuts()
        {
            if (_shortcutService == null || _flagService == null) return;
            var shortcuts = _shortcutService.CurrentSettings.Shortcuts;

            // 既存の Flag_ エントリを一旦削除
            var flagKeys = shortcuts.Keys.Where(k => k.StartsWith("Flag_")).ToList();
            foreach (var k in flagKeys) shortcuts.Remove(k);

            // 現在のフラグ定義からショートカットを登録
            foreach (var flag in _flagService.Flags)
            {
                if (!flag.HasShortcut) continue;
                shortcuts[$"Flag_{flag.Id}"] = new ShortcutKey
                {
                    Key = flag.ShortcutKey,
                    Modifiers = flag.ShortcutModifiers,
                    Description = $"フラグ: {flag.DisplayName}"
                };
            }

            _shortcutService.Save();
            Logger.Info("フラグショートカットを同期しました");
        }

        private static bool TryExecute(ICommand cmd)
        {
            if (cmd.CanExecute(null)) { cmd.Execute(null); return true; }
            return false;
        }

        private void MoveSheet(int direction)
        {
            if (SelectedTabIndex == 0)
                MoveInCollection(NormalSheet.NormalSheets, NormalSheet.SelectedNormalSheet, s => NormalSheet.SelectedNormalSheet = s, direction);
            else
                MoveInCollection(EastSheet.EastSheets, EastSheet.SelectedEastSheet, s => EastSheet.SelectedEastSheet = s, direction);
        }

        private static void MoveInCollection(ObservableCollection<string> list, string current, Action<string> setter, int direction)
        {
            if (list.Count == 0) return;
            int index = list.IndexOf(current) + direction;
            if (index < 0) index = list.Count - 1;
            else if (index >= list.Count) index = 0;
            setter(list[index]);
        }

        #endregion

        #region ウィンドウ管理

        private void OpenWindow<T>(string displayName) where T : Window, new()
        {
            try
            {
                Logger.Info($"{displayName}ウィンドウを開きます");
                new T { Owner = Application.Current.MainWindow }.Show();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"{displayName}ウィンドウを開く際にエラーが発生");
                MessageBox.Show($"ウィンドウを開けませんでした: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenSettings()
        {
            var vm = new SettingsWindowViewModel(Rates, _excelHandler, RatesFilePath, this, _shortcutService, _backupService, _flagService, _vehicleSettingsService);
            new Views.SettingsWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private void OpenPdfImport()
        {
            var apiKey = LoadApiKey();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                var inputDialog = new Views.ApiKeyInputWindow { Owner = Application.Current.MainWindow };
                if (inputDialog.ShowDialog() != true) return;
                apiKey = inputDialog.ApiKey;
                if (!string.IsNullOrWhiteSpace(apiKey))
                    SaveApiKey(apiKey);
            }

            var vm = new PdfImportViewModel(NormalSheet, Log, apiKey);
            new Views.PdfImportWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private void OpenHelp()
        {
            try
            {
                if (File.Exists(HelpFilePath))
                { Process.Start(new ProcessStartInfo(HelpFilePath) { UseShellExecute = true }); Log("ヘルプファイルを開きました。"); }
                else
                    MessageBox.Show("ヘルプファイル (readme.pdf) が見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ヘルプファイルを開けませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                Logger.Error(ex, "ヘルプファイルのオープン中にエラーが発生しました。");
            }
        }

        private void OpenEditWindow()
        {
            if (SelectedRow == null) return;
            var vm = new EditWindowViewModel(this, NormalSheet.SelectedNormalSheet, SelectedRow);
            new Views.EditWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        private void DeleteSelectedRow()
        {
            if (SelectedRow == null) return;
            var sheet = NormalSheet.SelectedNormalSheet;
            var rowIndex = SelectedRow.RowIndex;
            int idToDelete = (SelectedRow.DbId > 0) ? (int)SelectedRow.DbId : rowIndex;
            if (MessageBox.Show($"選択した行({rowIndex}行目)を削除しますか？\nこの操作は元に戻せません。",
                    "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                _excelHandler.DeleteRows(sheet, new List<int> { idToDelete });
                UpdatePreview();
                if (_dbService == null) _excelHandler.Save();
                Log($"[{sheet}] から {rowIndex}行目のデータを削除しました。");
            }
        }

        #endregion

        #region バージョン確認

        private async Task CheckForUpdate()
        {
            string currentVersion = "0.0.0";
            try
            {
                if (File.Exists(VersionFilePath))
                {
                    var versionData = JObject.Parse(await File.ReadAllTextAsync(VersionFilePath));
                    currentVersion = versionData["latest_version"]?.ToString() ?? "0.0.0";
                    AppVersion = $"v{currentVersion}";
                    Logger.Info($"ローカルバージョン: {currentVersion}");
                }
                else Logger.Warn($"version.json が見つかりません: {VersionFilePath}");
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "version.json の読み込みに失敗しました。バージョンチェックをスキップします。");
                return;
            }
            await new UpdateService(currentVersion, "", VersionInfoUrl, ReleasesPageUrl, Log).CheckForUpdateAsync();
        }

        #endregion

        #region ログ

        private void Log(string message)
        {
            Logger.Info(message);
            void Update()
            {
                _logBuilder.AppendLine(message);
                var lines = _logBuilder.ToString().Split('\n');
                if (lines.Length > MaxLogLines)
                {
                    _logBuilder.Clear();
                    _logBuilder.Append(string.Join("\n", lines.Skip(lines.Length - MaxLogLines)));
                }
                LogText = _logBuilder.ToString();
            }
            if (Application.Current.Dispatcher.CheckAccess()) Update();
            else Application.Current.Dispatcher.Invoke(Update);
        }

        #endregion
    }
}
