using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using Newtonsoft.Json;
using NLog;

namespace HansoInputTool.Services
{
    internal static class BranchNameResolver
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static List<BranchRule> _rules = new();
        private static bool _loaded = false;
        private const string RulesPath = "data/branch_rules.json";

        private static void EnsureLoaded()
        {
            if (_loaded) return;
            try
            {
                if (File.Exists(RulesPath))
                {
                    var json = File.ReadAllText(RulesPath);
                    var parsed = JsonConvert.DeserializeObject<List<BranchRule>>(json);
                    if (parsed != null && parsed.Count > 0) _rules = parsed;
                }
                else
                {
                    Logger.Warn($"Branch rules file not found: {RulesPath}. Using built-in fallbacks.");
                    _rules = new List<BranchRule>();
                }
                _rules = _rules.OrderBy(r => r.Priority).ToList();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Failed to load branch rules. Using built-in fallbacks.");
                _rules = new List<BranchRule>();
            }
            finally
            {
                _loaded = true;
            }
        }

        public static (string Branch, string Number) Resolve(string sheetName)
        {
            EnsureLoaded();

            foreach (var rule in _rules)
            {
                try
                {
                    var m = Regex.Match(sheetName, rule.Pattern);
                    if (!m.Success) continue;
                    string branch;
                    if (!string.IsNullOrEmpty(rule.BranchGroup) && m.Groups[rule.BranchGroup]?.Success == true)
                        branch = m.Groups[rule.BranchGroup].Value;
                    else if (m.Groups["branch"]?.Success == true)
                        branch = m.Groups["branch"].Value;
                    else
                        branch = m.Value;

                    string number = "";
                    if (!string.IsNullOrEmpty(rule.NumberGroup) && rule.NumberGroup != "null" && m.Groups[rule.NumberGroup]?.Success == true)
                        number = m.Groups[rule.NumberGroup].Value;
                    else if (m.Groups["number"]?.Success == true)
                        number = m.Groups["number"].Value;

                    return (branch?.Trim() ?? sheetName, number ?? "");
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, $"Branch rule threw for pattern: {rule.Pattern}");
                }
            }

            // フォールバック（既存のロジックを踏襲）
            if (sheetName.Contains("東日本セレモニー"))
            {
                var numberMatch = Regex.Match(sheetName, @"\d+$");
                return ("東日本", numberMatch.Success ? numberMatch.Value : "");
            }
            var parts = sheetName.Split(' ');
            if (parts.Length > 1 && int.TryParse(parts.Last(), out _))
            {
                return (string.Join(" ", parts.Take(parts.Length - 1)), parts.Last());
            }
            return (sheetName, "");
        }
    }
}
