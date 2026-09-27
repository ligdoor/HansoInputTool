using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using HansoInputTool.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace HansoInputTool.Services
{
    /// <summary>
    /// PDF解析で使用するAIプロバイダー設定を保存・読み込みします。
    /// APIキーはWindows DPAPIで暗号化して保存します。
    /// </summary>
    public class AiSettingsService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("HansoInputTool_AiSettings_v1");
        private readonly string _filePath;

        public AiSettingsService(string filePath)
        {
            _filePath = filePath;
        }

        public AiSettings Load()
        {
            try
            {
                if (!File.Exists(_filePath))
                    return new AiSettings();

                var obj = JObject.Parse(File.ReadAllText(_filePath));
                var settings = new AiSettings
                {
                    Provider = obj["provider"]?.ToString() ?? "Anthropic",
                    Model = obj["model"]?.ToString() ?? "claude-haiku-4-5-20251001",
                    Endpoint = obj["endpoint"]?.ToString()
                };

                var encrypted = obj["api_key"]?.ToString();
                if (!string.IsNullOrWhiteSpace(encrypted))
                {
                    var bytes = ProtectedData.Unprotect(
                        Convert.FromBase64String(encrypted), Entropy, DataProtectionScope.LocalMachine);
                    settings.ApiKey = Encoding.UTF8.GetString(bytes);
                }

                // 旧版Claude専用設定からの移行
                if (string.IsNullOrWhiteSpace(settings.ApiKey))
                {
                    var legacy = obj["claude_api_key"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(legacy))
                        settings.ApiKey = TryUnprotectLegacy(legacy);
                }

                return settings;
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "AI設定の読み込みに失敗しました。");
                return new AiSettings();
            }
        }

        public void Save(AiSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath));

            var obj = new JObject
            {
                ["provider"] = settings.Provider ?? "Anthropic",
                ["model"] = settings.Model ?? "",
                ["endpoint"] = settings.Endpoint ?? ""
            };

            if (!string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                var encrypted = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(settings.ApiKey), Entropy, DataProtectionScope.LocalMachine);
                obj["api_key"] = Convert.ToBase64String(encrypted);
            }

            File.WriteAllText(_filePath, obj.ToString(Formatting.Indented));
        }

        private static string TryUnprotectLegacy(string value)
        {
            try
            {
                // 旧版 v2 DPAPI
                var bytes = ProtectedData.Unprotect(
                    Convert.FromBase64String(value),
                    Encoding.UTF8.GetBytes("HansoInputTool_ApiKey_v2"),
                    DataProtectionScope.LocalMachine);
                return Encoding.UTF8.GetString(bytes);
            }
            catch { return null; }
        }
    }
}
