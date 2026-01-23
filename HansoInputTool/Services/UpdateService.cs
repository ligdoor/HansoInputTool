using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json.Linq;

namespace HansoInputTool.Services
{
    public class UpdateService
    {
        private readonly string _currentVersion;
        private readonly string _githubToken;
        private readonly string _versionInfoUrl;
        private readonly string _releasesPageUrl; // 固定のリリーススページURLを保持
        private readonly Action<string> _logAction;

        public UpdateService(string currentVersion, string githubToken, string versionInfoUrl, string releasesPageUrl, Action<string> logAction)
        {
            _currentVersion = currentVersion;
            _githubToken = githubToken;
            _versionInfoUrl = versionInfoUrl;
            _releasesPageUrl = releasesPageUrl; // コンストラクタで受け取る
            _logAction = logAction;
        }

        public async Task CheckForUpdateAsync()
        {
            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.ParseAdd("HansoInputTool");
                if (!string.IsNullOrEmpty(_githubToken))
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", _githubToken);
                }
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github.v3.raw"));

                var response = await client.GetStringAsync(_versionInfoUrl);
                var data = JObject.Parse(response);

                string latestVersionStr = data["latest_version"]?.ToString();

                if (Version.TryParse(latestVersionStr, out var latestVersion) &&
                    Version.TryParse(_currentVersion, out var currentVersion))
                {
                    if (latestVersion > currentVersion)
                    {
                        var result = MessageBox.Show(
                            $"新しいバージョン ({latestVersionStr}) が利用可能です。\n" +
                            $"現在のバージョン: {_currentVersion}\n\n" +
                            "ダウンロードページを開きますか？",
                            "アップデート通知",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information);

                        if (result == MessageBoxResult.Yes)
                        {
                            // version.jsonのURLではなく、固定のリリーススページURLを開く
                            Process.Start(new ProcessStartInfo(_releasesPageUrl) { UseShellExecute = true });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logAction?.Invoke($"[UPDATE CHECK FAILED]: {ex.Message}");
            }
        }
    }
}