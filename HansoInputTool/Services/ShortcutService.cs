using System;
using System.IO;
using HansoInputTool.Models;
using Newtonsoft.Json;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// ショートカット設定の読み書きを管理するサービス
    /// </summary>
    public class ShortcutService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly string _settingsFilePath;
        private ShortcutSettings _currentSettings;

        public ShortcutSettings CurrentSettings => _currentSettings;

        public ShortcutService(string settingsFilePath)
        {
            _settingsFilePath = settingsFilePath;
            Load();
        }

        /// <summary>
        /// 設定をファイルから読み込み
        /// </summary>
        public void Load()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    var json = File.ReadAllText(_settingsFilePath);
                    _currentSettings = JsonConvert.DeserializeObject<ShortcutSettings>(json);
                    
                    // 新しいショートカットが追加された場合のマージ処理
                    var defaults = ShortcutSettings.GetDefaultShortcuts();
                    foreach (var kvp in defaults)
                    {
                        if (!_currentSettings.Shortcuts.ContainsKey(kvp.Key))
                        {
                            _currentSettings.Shortcuts[kvp.Key] = kvp.Value;
                        }
                    }
                    
                    Logger.Info("ショートカット設定を読み込みました。");
                }
                else
                {
                    _currentSettings = new ShortcutSettings();
                    Save(); // デフォルト設定を保存
                    Logger.Info("ショートカット設定ファイルが存在しないため、デフォルト設定を作成しました。");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "ショートカット設定の読み込み中にエラーが発生しました。デフォルト設定を使用します。");
                _currentSettings = new ShortcutSettings();
            }
        }

        /// <summary>
        /// 設定をファイルに保存
        /// </summary>
        public void Save()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_currentSettings, Formatting.Indented);
                File.WriteAllText(_settingsFilePath, json);
                Logger.Info("ショートカット設定を保存しました。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "ショートカット設定の保存中にエラーが発生しました。");
                throw;
            }
        }

        /// <summary>
        /// ショートカット設定を更新
        /// </summary>
        public void UpdateSettings(ShortcutSettings newSettings)
        {
            _currentSettings = newSettings;
        }

        /// <summary>
        /// 特定のショートカットを取得
        /// </summary>
        public ShortcutKey GetShortcut(string actionName)
        {
            if (_currentSettings.Shortcuts.TryGetValue(actionName, out var shortcut))
            {
                return shortcut;
            }
            return null;
        }

        /// <summary>
        /// デフォルト設定にリセット
        /// </summary>
        public void ResetToDefaults()
        {
            _currentSettings = new ShortcutSettings();
            Save();
            Logger.Info("ショートカット設定をデフォルトにリセットしました。");
        }
    }
}
