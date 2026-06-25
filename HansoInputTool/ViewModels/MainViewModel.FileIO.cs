using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using HansoInputTool.Models;
using HansoInputTool.Services;
using Microsoft.Win32;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;

namespace HansoInputTool.ViewModels
{
    public partial class MainViewModel
    {
        #region ファイル操作

        private void LoadGeppoFile()
        {
            var dialog = new OpenFileDialog { Title = "編集する実績月報ファイルを選択", Filter = "Excel ファイル (*.xlsx)|*.xlsx" };
            if (dialog.ShowDialog() != true) return;
            if (MessageBox.Show("選択したファイルの内容で現在の作業内容を上書きします。\nよろしいですか？",
                    "上書き確認", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.Cancel) return;
            try
            {
                File.Copy(dialog.FileName, InputFilePath, true);
                _excelHandler.Load();

                // ファイル名から期・月・R年を解析してUIに反映
                // 例: "46期 4月 R7 アルス搬送・霊柩車　実績月報.xlsx"
                var fileName = Path.GetFileNameWithoutExtension(dialog.FileName);
                var m = Regex.Match(fileName, @"(\d+)期.*?(\d+)月.*?[RＲ](\d+)");
                if (m.Success)
                {
                    Period  = m.Groups[1].Value;
                    Month   = m.Groups[2].Value;
                    RNumber = m.Groups[3].Value;
                    Log($"ファイル名から期・月・R年を読み込みました: {Period}期 {Month}月 R{RNumber}");
                }
                else
                {
                    // [No.5修正] ファイル名パースに失敗した場合、画面上部に入力済みの
                    // 期・月・R年をそのまま使う。未入力なら警告してインポートを中止する。
                    Logger.Warn($"ファイル名から期・月・R年を解析できませんでした: {fileName}");
                    Log("ファイル名から期・月・R年を読み取れませんでした。画面上部に入力された値を使用します。");
                }

                // [No.5修正] パース成否にかかわらず必ずセッションを確定してからインポートする
                if (_dbService != null)
                {
                    if (string.IsNullOrWhiteSpace(Period) || string.IsNullOrWhiteSpace(Month) || string.IsNullOrWhiteSpace(RNumber))
                    {
                        MessageBox.Show("期・月・R年が設定されていません。\n画面上部の入力欄を入力してから再度読み込んでください。",
                            "入力不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    _dbService.GetOrCreateSession(Period, Month, RNumber);
                    Log($"月データセッション準備完了: {Period}期 {Month}月 R{RNumber}");
                }

                // DB使用時: セッション確定後にExcelデータをDBにインポート
                if (_dbService != null)
                {
                    _excelHandler.ImportFromExcelToDb(_dbService, _flagService);
                    Log($"[DB] 通常系シートのデータをDBにインポートしました。");
                }

                ReloadAllData();

                Log($"実績月報 '{Path.GetFileName(dialog.FileName)}' を読み込みました。");
                MessageBox.Show("実績月報のデータを読み込みました。", "読み込み完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "実績月報ファイルの読み込み中にエラーが発生しました。");
                MessageBox.Show("ファイルの読み込みに失敗しました。\n詳細はログファイルを確認してください。",
                    "読み込みエラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveInputFile()
        {
            try
            {
                _excelHandler.Save();
                MessageBox.Show("現在の入力内容を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                Log("--- 入力内容を保存しました ---");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "入力内容の保存中にエラーが発生しました。");
                MessageBox.Show("保存に失敗しました。\n詳細はログファイルを確認してください。",
                    "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region バックアップ

        private void CreateManualBackup()
        {
            try
            {
                var inputBackup    = _backupService.CreateManualBackup(InputFilePath,    "手動保存");
                var templateBackup = _backupService.CreateManualBackup(TemplateFilePath, "手動保存");
                if (inputBackup != null && templateBackup != null)
                {
                    MessageBox.Show(
                        $"バックアップを作成しました。\n\nInput.xlsx: {Path.GetFileName(inputBackup)}\nTemplate.xlsx: {Path.GetFileName(templateBackup)}\n\n保存場所: backupsフォルダ",
                        "バックアップ完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    Log("手動バックアップを作成しました。");
                }
                else
                {
                    MessageBox.Show("バックアップの作成に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "手動バックアップの作成中にエラーが発生しました");
                MessageBox.Show("バックアップの作成に失敗しました。\n詳細はログファイルを確認してください。",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenRestoreBackupWindow()
        {
            var vm = new RestoreBackupWindowViewModel(_backupService, InputFilePath, TemplateFilePath, this);
            new Views.RestoreBackupWindow(vm) { Owner = Application.Current.MainWindow }.ShowDialog();
        }

        #endregion

        #region APIキー管理

        private static readonly string ApiSettingsFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "data", "api_settings.json");

        // [No.9修正] AESキーのハードコードを廃止し、WindowsのDPAPI（ProtectedData）に移行。
        // DPAPIはログオン中のWindowsユーザー資格情報を使って暗号化するため、
        // ソースコードやバイナリを入手しても復号できない。
        // 共有サーバー環境では DataProtectionScope.LocalMachine を使い、
        // 同一マシン上のどのユーザーでも復号できるようにする。
        private static readonly byte[] DpapiEntropy =
            System.Text.Encoding.UTF8.GetBytes("HansoInputTool_ApiKey_v2");

        private string LoadApiKey()
        {
            try
            {
                if (!File.Exists(ApiSettingsFilePath)) return null;
                var json = File.ReadAllText(ApiSettingsFilePath);
                var obj = JObject.Parse(json);
                var encrypted = obj["claude_api_key"]?.ToString();
                if (string.IsNullOrEmpty(encrypted)) return null;

                // 旧AES形式（v1）からの移行: DPAPIで復号を試み、失敗したら旧形式を試す
                try
                {
                    var encryptedBytes = Convert.FromBase64String(encrypted);
                    var decryptedBytes = System.Security.Cryptography.ProtectedData.Unprotect(
                        encryptedBytes, DpapiEntropy,
                        System.Security.Cryptography.DataProtectionScope.LocalMachine);
                    return System.Text.Encoding.UTF8.GetString(decryptedBytes);
                }
                catch
                {
                    // 旧AES形式（v1）は読み取りのみサポート（保存時は自動でDPAPIに移行される）
                    Logger.Info("APIキー: 旧AES形式を検出。次回保存時にDPAPI形式に自動移行します。");
                    return LoadApiKeyLegacyAes(encrypted);
                }
            }
            catch { return null; }
        }

        /// <summary>旧AES形式（v1）のAPIキーを復号する（移行期間中のみ使用）</summary>
        private static string LoadApiKeyLegacyAes(string encryptedBase64)
        {
            try
            {
                var legacyKey = System.Text.Encoding.UTF8.GetBytes("HansoTool!AES256Key#2025$Secure!");
                var legacyIv  = System.Text.Encoding.UTF8.GetBytes("HansoIV!16Bytes!");
                using var aes = System.Security.Cryptography.Aes.Create();
                aes.Key = legacyKey;
                aes.IV  = legacyIv;
                using var decryptor = aes.CreateDecryptor();
                var encryptedBytes = Convert.FromBase64String(encryptedBase64);
                using var ms = new MemoryStream(encryptedBytes);
                using var cs = new System.Security.Cryptography.CryptoStream(
                    ms, decryptor, System.Security.Cryptography.CryptoStreamMode.Read);
                using var reader = new StreamReader(cs);
                return reader.ReadToEnd();
            }
            catch { return null; }
        }

        private void SaveApiKey(string apiKey)
        {
            try
            {
                // DPAPI（LocalMachine スコープ）で暗号化して保存
                var plainBytes = System.Text.Encoding.UTF8.GetBytes(apiKey);
                var encryptedBytes = System.Security.Cryptography.ProtectedData.Protect(
                    plainBytes, DpapiEntropy,
                    System.Security.Cryptography.DataProtectionScope.LocalMachine);
                var encrypted = Convert.ToBase64String(encryptedBytes);

                var obj = new JObject { ["claude_api_key"] = encrypted };
                File.WriteAllText(ApiSettingsFilePath, obj.ToString());
                Logger.Info("APIキーをDPAPI形式で保存しました。");
            }
            catch (Exception ex) { Logger.Warn(ex, "APIキーの保存に失敗しました"); }
        }

        #endregion

        #region COM フォルダ選択ダイアログ

        private string ShowFolderBrowserDialog(string title)
        {
            var dialog = (IFileOpenDialog_MV)new FileOpenDialog_MV();
            try
            {
                dialog.SetOptions(FOS_PICKFOLDERS_MV | FOS_FORCEFILESYSTEM_MV);
                dialog.SetTitle(title);
                int hr = dialog.Show(IntPtr.Zero);
                if (hr < 0) return null;
                dialog.GetResult(out IShellItem_MV item);
                item.GetDisplayName(SIGDN_FILESYSPATH_MV, out string path);
                return path;
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }

        private const uint FOS_PICKFOLDERS_MV    = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM_MV = 0x00000040;
        private const uint SIGDN_FILESYSPATH_MV   = 0x80058000;

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialog_MV { }

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog_MV
        {
            [PreserveSig] int Show(IntPtr parent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem_MV psi);
            void SetFolder(IShellItem_MV psi);
            void GetFolder(out IShellItem_MV ppsi);
            void GetCurrentSelection(out IShellItem_MV ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem_MV ppsi);
            void AddPlace(IShellItem_MV psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppsai);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem_MV
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem_MV ppsi);
            void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem_MV psi, uint hint, out int piOrder);
        }

        #endregion
    }
}
