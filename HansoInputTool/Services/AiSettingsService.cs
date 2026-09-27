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
    /// APIキーはWindows DPAPIで暗号化し、AI設定管理パスワードはPBKDF2ハッシュで保存します。
    /// </summary>
    public class AiSettingsService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("HansoInputTool_AiSettings_v1");
        private const int PasswordIterations = 120_000;
        private const int SaltSize = 16;
        private const int HashSize = 32;
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
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            var obj = LoadRawObject();
            obj["provider"] = settings.Provider ?? "Anthropic";
            obj["model"] = settings.Model ?? "";
            obj["endpoint"] = settings.Endpoint ?? "";

            if (!string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                var encrypted = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(settings.ApiKey), Entropy, DataProtectionScope.LocalMachine);
                obj["api_key"] = Convert.ToBase64String(encrypted);
            }
            else
            {
                obj.Remove("api_key");
            }

            // 旧形式のキーは保存時に残さない
            obj.Remove("claude_api_key");
            WriteRawObject(obj);
        }

        public bool HasAdminPassword()
        {
            try
            {
                var obj = LoadRawObject();
                return !string.IsNullOrWhiteSpace(obj["admin_password_hash"]?.ToString())
                    && !string.IsNullOrWhiteSpace(obj["admin_password_salt"]?.ToString());
            }
            catch
            {
                return false;
            }
        }

        public void SetAdminPassword(string password)
        {
            if (string.IsNullOrWhiteSpace(password) || password.Length < 6)
                throw new ArgumentException("AI設定管理パスワードは6文字以上で設定してください。", nameof(password));

            var salt = RandomNumberGenerator.GetBytes(SaltSize);
            var hash = DerivePasswordHash(password, salt);
            var obj = LoadRawObject();
            obj["admin_password_salt"] = Convert.ToBase64String(salt);
            obj["admin_password_hash"] = Convert.ToBase64String(hash);
            obj["admin_password_iterations"] = PasswordIterations;
            WriteRawObject(obj);
        }

        public bool VerifyAdminPassword(string password)
        {
            if (string.IsNullOrEmpty(password)) return false;

            try
            {
                var obj = LoadRawObject();
                var saltText = obj["admin_password_salt"]?.ToString();
                var hashText = obj["admin_password_hash"]?.ToString();
                if (string.IsNullOrWhiteSpace(saltText) || string.IsNullOrWhiteSpace(hashText))
                    return false;

                var iterations = obj["admin_password_iterations"]?.Value<int>() ?? PasswordIterations;
                if (iterations < 50_000 || iterations > 1_000_000)
                    iterations = PasswordIterations;

                var salt = Convert.FromBase64String(saltText);
                var expected = Convert.FromBase64String(hashText);
                var actual = DerivePasswordHash(password, salt, iterations);
                return CryptographicOperations.FixedTimeEquals(actual, expected);
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "AI設定管理パスワードの検証に失敗しました。");
                return false;
            }
        }

        private static byte[] DerivePasswordHash(string password, byte[] salt, int iterations = PasswordIterations)
        {
            return Rfc2898DeriveBytes.Pbkdf2(
                password,
                salt,
                iterations,
                HashAlgorithmName.SHA256,
                HashSize);
        }

        private JObject LoadRawObject()
        {
            if (!File.Exists(_filePath))
                return new JObject();

            try
            {
                return JObject.Parse(File.ReadAllText(_filePath));
            }
            catch
            {
                return new JObject();
            }
        }

        private void WriteRawObject(JObject obj)
        {
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);
            File.WriteAllText(_filePath, obj.ToString(Formatting.Indented));
        }

        private static string TryUnprotectLegacy(string value)
        {
            try
            {
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
