using System.Collections.Generic;
using System.Linq;
using HansoInputTool.Models;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 「この車両シートはどのタブ（＝どの入力方式）か」を判定する窓口。
    /// これまで各所に散らばっていた「シート名に『東日本』を含むか」の判定を、ここに集約していく。
    ///
    /// 判定の優先順位:
    ///   1) vehicle_settings.json の TabId（存在するタブを指している場合）
    ///   2) 従来ルール（名前に「東日本」を含む → 月合計入力、それ以外 → 日別入力）
    /// 2) があるため、既存データは何も書き換えなくても今までどおりに振り分けられる。
    /// </summary>
    public class SheetTabResolver
    {
        private const string LegacyTotalKeyword = "東日本";

        /// <summary>従来から固定で扱っている営業所名の並び順（出力順の基準）</summary>
        public static readonly string[] LegacyBranchOrder = { "CH富士吉田", "CH大月", "CH東富士", "東日本セレモニー" };

        /// <summary>
        /// アプリ全体で共有する判定窓口（MainViewModel が起動時に設定する）。
        /// 年度集計画面のように、メイン画面から直接受け渡しできない場所で使う。
        /// </summary>
        public static SheetTabResolver Current { get; set; }

        private readonly TabSettingsService _tabs;
        private readonly VehicleSettingsService _vehicles;

        public SheetTabResolver(TabSettingsService tabs, VehicleSettingsService vehicles)
        {
            _tabs = tabs;
            _vehicles = vehicles;
        }

        /// <summary>シートが所属するタブを返す。該当タブが見つからない場合は null。</summary>
        public TabDefinition GetTab(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) return null;

            if (_vehicles?.Settings != null
                && _vehicles.Settings.TryGetValue(sheetName, out var cfg)
                && !string.IsNullOrEmpty(cfg?.TabId))
            {
                var assigned = _tabs.GetTab(cfg.TabId);
                if (assigned != null) return assigned;
            }

            // 従来ルール
            bool legacyTotal = sheetName.Contains(LegacyTotalKeyword);
            var legacyTab = _tabs.GetTab(legacyTotal ? DefaultTabIds.Total : DefaultTabIds.Daily);
            return legacyTab ?? _tabs.FindFirstByMode(legacyTotal ? TabInputMode.Total : TabInputMode.Daily);
        }

        /// <summary>シートの入力方式を返す（タブが見つからなくても従来ルールで必ず答える）</summary>
        public string GetInputMode(string sheetName)
        {
            var tab = GetTab(sheetName);
            if (tab != null) return tab.InputMode;
            return !string.IsNullOrEmpty(sheetName) && sheetName.Contains(LegacyTotalKeyword)
                ? TabInputMode.Total
                : TabInputMode.Daily;
        }

        /// <summary>月合計入力のシートか（従来の「東日本シートか」に相当）</summary>
        public bool IsTotalMode(string sheetName) => GetInputMode(sheetName) == TabInputMode.Total;

        // ── 出力順・グループ名（実績月報／月間集計／年度集計で共通に使う） ──

        private bool IsDefaultTab(TabDefinition tab)
            => tab != null && (tab.Id == DefaultTabIds.Daily || tab.Id == DefaultTabIds.Total);

        private int IndexOfTab(string tabId)
        {
            var list = _tabs.Tabs;
            for (int i = 0; i < list.Count; i++)
                if (list[i].Id == tabId) return i;
            return list.Count;
        }

        /// <summary>シートが所属するタブの並び順（0始まり）。タブが判定できない場合はタブ数（＝最後）</summary>
        public int GetTabOrder(string sheetName)
        {
            var tab = GetTab(sheetName);
            return tab == null ? _tabs.Tabs.Count : IndexOfTab(tab.Id);
        }

        /// <summary>
        /// 利用者が作ったタブ（既定の「日別入力」「月合計入力」以外）に所属する車両の、グループ名（＝タブ名）。
        /// 該当しない場合は null。
        /// </summary>
        public string GetCustomGroupLabel(string sheetName)
        {
            var tab = GetTab(sheetName);
            return IsDefaultTab(tab) ? null : tab?.Name;
        }

        /// <summary>
        /// 年度集計用：利用者が作ったタブの車両について、グループ名（タブ名）と車番（末尾の数字。無ければシート名）を返す。
        /// </summary>
        public bool TryGetCustomGroup(string sheetName, out string groupLabel, out string vehicleNo)
        {
            groupLabel = GetCustomGroupLabel(sheetName);
            vehicleNo = null;
            if (groupLabel == null) return false;

            var m = System.Text.RegularExpressions.Regex.Match(sheetName, @"\d+$");
            vehicleNo = m.Success ? m.Value : sheetName;
            return true;
        }

        /// <summary>
        /// グループ名（営業所名／タブ名）の並び順。タブの並び順を優先し、同じタブ内では従来の営業所順になる。
        /// 不明なグループは int.MaxValue。
        /// </summary>
        public int GetGroupOrder(string label)
        {
            if (string.IsNullOrEmpty(label)) return int.MaxValue;

            int legacy = System.Array.IndexOf(LegacyBranchOrder, label);
            if (legacy >= 0)
            {
                // 従来の営業所は、所属タブ（通常は既定タブ）の位置で並べる
                string representative = label == "東日本セレモニー" ? "東日本セレモニー 1" : label + " 寝台車 1";
                return GetTabOrder(representative) * 10 + legacy;
            }

            var list = _tabs.Tabs;
            for (int i = 0; i < list.Count; i++)
                if (!IsDefaultTab(list[i]) && list[i].Name == label) return i * 10 + 9;

            return int.MaxValue;
        }

        /// <summary>指定タブに所属するシート名だけを、元の並びのまま返す</summary>
        public List<string> GetSheetsForTab(string tabId, IEnumerable<string> allSheets)
            => allSheets.Where(s => GetTab(s)?.Id == tabId).ToList();
    }
}
