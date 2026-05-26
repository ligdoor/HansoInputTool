using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using HansoInputTool.Models;
using HansoInputTool.ViewModels.Base;

namespace HansoInputTool.ViewModels
{
    /// <summary>
    /// ショートカット設定画面のViewModel
    /// </summary>
    public class ShortcutSettingItem : ObservableObject
    {
        public string ActionName { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }

        private Key _key;
        public Key Key
        {
            get => _key;
            set
            {
                if (SetProperty(ref _key, value))
                {
                    OnPropertyChanged(nameof(DisplayString));
                }
            }
        }

        private ModifierKeys _modifiers;
        public ModifierKeys Modifiers
        {
            get => _modifiers;
            set
            {
                if (SetProperty(ref _modifiers, value))
                {
                    OnPropertyChanged(nameof(DisplayString));
                }
            }
        }

        private bool _isRecording;
        public bool IsRecording
        {
            get => _isRecording;
            set => SetProperty(ref _isRecording, value);
        }

        public string DisplayString
        {
            get
            {
                if (Key == Key.None)
                    return "(未設定)";

                var parts = new System.Collections.Generic.List<string>();

                if (Modifiers.HasFlag(ModifierKeys.Control))
                    parts.Add("Ctrl");
                if (Modifiers.HasFlag(ModifierKeys.Alt))
                    parts.Add("Alt");
                if (Modifiers.HasFlag(ModifierKeys.Shift))
                    parts.Add("Shift");

                parts.Add(GetKeyDisplayName(Key));

                return string.Join("+", parts);
            }
        }

        private static string GetKeyDisplayName(Key key)
        {
            return key switch
            {
                Key.OemComma => ",",
                Key.OemPeriod => ".",
                Key.OemPlus => "+",
                Key.OemMinus => "-",
                Key.OemQuestion => "/",
                Key.OemOpenBrackets => "[",
                Key.OemCloseBrackets => "]",
                Key.OemSemicolon => ";",
                Key.OemQuotes => "'",
                Key.Left => "←",
                Key.Right => "→",
                Key.Up => "↑",
                Key.Down => "↓",
                Key.Enter => "Enter",
                Key.Tab => "Tab",
                Key.Delete => "Delete",
                Key.Back => "Backspace",
                Key.Escape => "Esc",
                Key.Space => "Space",
                _ => key.ToString()
            };
        }

        /// <summary>
        /// ShortcutKeyオブジェクトに変換
        /// </summary>
        public ShortcutKey ToShortcutKey()
        {
            return new ShortcutKey
            {
                Key = Key,
                Modifiers = Modifiers,
                Description = Description
            };
        }

        /// <summary>
        /// ShortcutKeyオブジェクトから作成
        /// </summary>
        public static ShortcutSettingItem FromShortcutKey(string actionName, ShortcutKey shortcutKey)
        {
            return new ShortcutSettingItem
            {
                ActionName  = actionName,
                // Flag_xxx はDescriptionに「フラグ: 表示名」が入っているのでそちらを優先
                DisplayName = actionName.StartsWith("Flag_")
                    ? (shortcutKey.Description ?? ShortcutSettings.GetActionDisplayName(actionName))
                    : ShortcutSettings.GetActionDisplayName(actionName),
                Description = shortcutKey.Description,
                Key = shortcutKey.Key,
                Modifiers = shortcutKey.Modifiers
            };
        }
    }

    /// <summary>
    /// ショートカット設定のコレクションを管理
    /// </summary>
    public class ShortcutSettingsViewModel : ObservableObject
    {
        public ObservableCollection<ShortcutSettingItem> ShortcutItems { get; }

        private ShortcutSettingItem _selectedItem;
        public ShortcutSettingItem SelectedItem
        {
            get => _selectedItem;
            set => SetProperty(ref _selectedItem, value);
        }

        public ICommand ResetToDefaultsCommand { get; }
        public ICommand ClearShortcutCommand { get; }

        public ShortcutSettingsViewModel(ShortcutSettings settings)
        {
            ShortcutItems = new ObservableCollection<ShortcutSettingItem>();
            
            foreach (var kvp in settings.Shortcuts.OrderBy(x => x.Key))
            {
                ShortcutItems.Add(ShortcutSettingItem.FromShortcutKey(kvp.Key, kvp.Value));
            }

            ResetToDefaultsCommand = new RelayCommand(_ => ResetToDefaults());
            ClearShortcutCommand = new RelayCommand(_ => ClearSelectedShortcut(), _ => SelectedItem != null);
        }

        private void ResetToDefaults()
        {
            var defaults = ShortcutSettings.GetDefaultShortcuts();
            foreach (var item in ShortcutItems)
            {
                if (defaults.TryGetValue(item.ActionName, out var defaultKey))
                {
                    item.Key = defaultKey.Key;
                    item.Modifiers = defaultKey.Modifiers;
                }
            }
        }

        private void ClearSelectedShortcut()
        {
            if (SelectedItem != null)
            {
                SelectedItem.Key = Key.None;
                SelectedItem.Modifiers = ModifierKeys.None;
            }
        }

        /// <summary>
        /// ShortcutSettingsオブジェクトに変換
        /// </summary>
        public ShortcutSettings ToShortcutSettings()
        {
            var settings = new ShortcutSettings
            {
                Shortcuts = ShortcutItems.ToDictionary(
                    item => item.ActionName,
                    item => item.ToShortcutKey()
                )
            };
            return settings;
        }

        /// <summary>
        /// 重複チェック
        /// </summary>
        public bool HasDuplicates(out string duplicateInfo)
        {
            var grouped = ShortcutItems
                .Where(x => x.Key != Key.None)
                .GroupBy(x => new { x.Key, x.Modifiers })
                .Where(g => g.Count() > 1)
                .ToList();

            if (grouped.Any())
            {
                var first = grouped.First();
                var names = string.Join(", ", first.Select(x => x.DisplayName));
                duplicateInfo = $"同じショートカット ({first.First().DisplayString}) が複数のアクションに設定されています: {names}";
                return true;
            }

            duplicateInfo = null;
            return false;
        }
    }
}
