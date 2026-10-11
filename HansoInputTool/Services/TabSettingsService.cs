using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HansoInputTool.Models;
using Newtonsoft.Json;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 入力タブの定義（tab_settings.json）を読み書きするサービス。
    /// ファイルが無い場合は、従来の2タブ（日別入力・月合計入力）を自動で作成する。
    /// </summary>
    public class TabSettingsService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly string _filePath;
        private List<TabDefinition> _tabs;

        public TabSettingsService(string filePath)
        {
            _filePath = filePath;
            _tabs = Load();
        }

        /// <summary>タブ一覧（並び順どおり）</summary>
        public IReadOnlyList<TabDefinition> Tabs => _tabs;

        public TabDefinition GetTab(string id)
            => string.IsNullOrEmpty(id) ? null : _tabs.FirstOrDefault(t => t.Id == id);

        /// <summary>指定した入力方式を持つ最初のタブを返す（無ければ null）</summary>
        public TabDefinition FindFirstByMode(string mode)
            => _tabs.FirstOrDefault(t => t.InputMode == mode);

        public static List<TabDefinition> CreateDefaults() => new()
        {
            new TabDefinition { Id = DefaultTabIds.Daily, Name = TabInputMode.DisplayName(TabInputMode.Daily), InputMode = TabInputMode.Daily },
            new TabDefinition { Id = DefaultTabIds.Total, Name = TabInputMode.DisplayName(TabInputMode.Total), InputMode = TabInputMode.Total },
        };

        private List<TabDefinition> Load()
        {
            try
            {
                if (!File.Exists(_filePath))
                {
                    var defaults = CreateDefaults();
                    WriteFile(defaults);
                    Logger.Info("tab_settings.json が無かったため、初期タブ（日別入力・月合計入力）を作成しました。");
                    return defaults;
                }

                var json = File.ReadAllText(_filePath);
                var loaded = JsonConvert.DeserializeObject<List<TabDefinition>>(json);
                if (IsValid(loaded, out var reason))
                {
                    Logger.Info($"tab_settings.json を読み込みました ({loaded.Count}件)");
                    return loaded;
                }

                // 壊れたファイルは上書きしない（手で直せるように残し、今回はメモリ上の初期値で動かす）
                Logger.Warn($"tab_settings.json の内容が不正です（{reason}）。初期タブで動作します。ファイルは変更しません。");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "tab_settings.json の読み込みに失敗しました。初期タブで動作します。");
            }
            return CreateDefaults();
        }

        private static bool IsValid(List<TabDefinition> tabs, out string reason)
        {
            reason = null;
            if (tabs == null || tabs.Count == 0) { reason = "タブが1つもありません"; return false; }
            if (tabs.Any(t => t == null || string.IsNullOrWhiteSpace(t.Id))) { reason = "IDが空のタブがあります"; return false; }
            if (tabs.Select(t => t.Id).Distinct().Count() != tabs.Count) { reason = "IDが重複しています"; return false; }
            if (tabs.Any(t => string.IsNullOrWhiteSpace(t.Name))) { reason = "名前が空のタブがあります"; return false; }
            if (tabs.Any(t => !TabInputMode.IsValid(t.InputMode))) { reason = "入力方式が不正なタブがあります"; return false; }
            return true;
        }

        /// <summary>タブ一覧を保存する（リストの順番がそのままタブの並び順になる）</summary>
        public void Save(IEnumerable<TabDefinition> tabs)
        {
            var list = tabs.ToList();
            if (!IsValid(list, out var reason))
                throw new InvalidOperationException($"タブ設定を保存できません: {reason}");

            WriteFile(list);
            _tabs = list;
            Logger.Info($"tab_settings.json を保存しました ({list.Count}件)");
        }

        public void Reload() => _tabs = Load();

        private void WriteFile(List<TabDefinition> tabs)
        {
            File.WriteAllText(_filePath, JsonConvert.SerializeObject(tabs, Formatting.Indented));
        }
    }
}
