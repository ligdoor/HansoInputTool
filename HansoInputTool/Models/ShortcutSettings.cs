using System;
using System.Collections.Generic;
using System.Windows.Input;
using Newtonsoft.Json;

namespace HansoInputTool.Models
{
    /// <summary>
    /// ショートカットキー設定を保持するクラス
    /// </summary>
    public class ShortcutSettings
    {
        /// <summary>
        /// ショートカットの一覧（キー: アクション名、値: キー設定）
        /// </summary>
        public Dictionary<string, ShortcutKey> Shortcuts { get; set; }

        public ShortcutSettings()
        {
            // デフォルト値で初期化
            Shortcuts = GetDefaultShortcuts();
        }

        /// <summary>
        /// デフォルトのショートカット設定を取得
        /// </summary>
        public static Dictionary<string, ShortcutKey> GetDefaultShortcuts()
        {
            return new Dictionary<string, ShortcutKey>
            {
                { "Save", new ShortcutKey { Key = Key.S, Modifiers = ModifierKeys.Control, Description = "入力内容を保存" } },
                { "Register", new ShortcutKey { Key = Key.Enter, Modifiers = ModifierKeys.Control, Description = "データを登録" } },
                { "NextSheet", new ShortcutKey { Key = Key.Right, Modifiers = ModifierKeys.Control, Description = "次のシートへ移動" } },
                { "PrevSheet", new ShortcutKey { Key = Key.Left, Modifiers = ModifierKeys.Control, Description = "前のシートへ移動" } },
                { "Transfer", new ShortcutKey { Key = Key.T, Modifiers = ModifierKeys.Control, Description = "転記を実行" } },
                { "OpenSettings", new ShortcutKey { Key = Key.OemComma, Modifiers = ModifierKeys.Control, Description = "設定を開く" } },
                { "SwitchTab", new ShortcutKey { Key = Key.Tab, Modifiers = ModifierKeys.Control, Description = "タブを切り替え" } },
                { "EditRow", new ShortcutKey { Key = Key.E, Modifiers = ModifierKeys.Control, Description = "選択行を編集" } },
                { "DeleteRow", new ShortcutKey { Key = Key.Delete, Modifiers = ModifierKeys.Control, Description = "選択行を削除" } },
                { "CreateBackup", new ShortcutKey { Key = Key.B, Modifiers = ModifierKeys.Control, Description = "バックアップを作成" } }
            };
        }

        /// <summary>
        /// アクション名の日本語表示を取得
        /// </summary>
        public static string GetActionDisplayName(string actionName)
        {
            return actionName switch
            {
                "Save" => "保存",
                "Register" => "登録",
                "NextSheet" => "次のシート",
                "PrevSheet" => "前のシート",
                "Transfer" => "転記実行",
                "OpenSettings" => "設定を開く",
                "SwitchTab" => "タブ切替",
                "EditRow" => "行を編集",
                "DeleteRow" => "行を削除",
                "CreateBackup" => "バックアップ作成",
                // Flag_xxx 形式はそのまま渡すとIDになるため
                // 呼び出し元でDescriptionを使うためここではプレフィックスだけ除去
                _ when actionName.StartsWith("Flag_") => actionName.Substring(5),
                _ => actionName
            };
        }
    }

    /// <summary>
    /// 個別のショートカットキー設定
    /// </summary>
    public class ShortcutKey
    {
        [JsonIgnore]
        public Key Key { get; set; }

        [JsonIgnore]
        public ModifierKeys Modifiers { get; set; }

        /// <summary>
        /// ショートカットの説明
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// JSON保存用：キー文字列
        /// </summary>
        [JsonProperty("Key")]
        public string KeyString
        {
            get => Key.ToString();
            set => Key = Enum.TryParse<Key>(value, out var k) ? k : Key.None;
        }

        /// <summary>
        /// JSON保存用：修飾キー文字列
        /// </summary>
        [JsonProperty("Modifiers")]
        public string ModifiersString
        {
            get => Modifiers.ToString();
            set => Modifiers = Enum.TryParse<ModifierKeys>(value, out var m) ? m : ModifierKeys.None;
        }

        /// <summary>
        /// 表示用の文字列（例: "Ctrl+S"）
        /// </summary>
        [JsonIgnore]
        public string DisplayString
        {
            get
            {
                var parts = new List<string>();

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

        /// <summary>
        /// キーの表示名を取得
        /// </summary>
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
        /// 指定されたキー入力がこのショートカットに一致するか
        /// </summary>
        public bool Matches(Key pressedKey, ModifierKeys pressedModifiers)
        {
            return Key == pressedKey && Modifiers == pressedModifiers;
        }
    }
}
