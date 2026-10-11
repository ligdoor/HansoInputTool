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

        /// <summary>指定タブに所属するシート名だけを、元の並びのまま返す</summary>
        public List<string> GetSheetsForTab(string tabId, IEnumerable<string> allSheets)
            => allSheets.Where(s => GetTab(s)?.Id == tabId).ToList();
    }
}
