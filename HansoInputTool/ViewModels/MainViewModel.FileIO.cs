using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using HansoInputTool.Models;
using HansoInputTool.Services;
using Microsoft.Win32;
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

                // DB使用時: 読み込んだExcelの通常系データをDBにインポート
                if (_dbService != null)
                {
                    _excelHandler.ImportFromExcelToDb(_dbService, _flagService);
                    Log($"[DB] 通常系シートのデータをDBにインポートしました。");
                }

                ReloadAllData();

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

                    // セッションを自動作成または既存に切替
                    if (_dbService != null)
                    {
                        _dbService.GetOrCreateSession(Period, Month, RNumber);
                        Log($"月データセッション準備完了: {Period}期 {Month}月 R{RNumber}");
                    }
                }

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

        // AES暗号化用の固定キー（変更しないこと）
        private static readonly byte[] AesKey = System.Text.Encoding.UTF8.GetBytes("HansoTool!AES256Key#2025$Secure!"); // 32バイト
        private static readonly byte[] AesIv  = System.Text.Encoding.UTF8.GetBytes("HansoIV!16Bytes!");                 // 16バイト

        private string LoadApiKey()
        {
            try
            {
                if (!File.Exists(ApiSettingsFilePath)) return null;
                var json = File.ReadAllText(ApiSettingsFilePath);
                var obj = JObject.Parse(json);
                var encrypted = obj["claude_api_key"]?.ToString();
                if (string.IsNullOrEmpty(encrypted)) return null;

                using var aes = System.Security.Cryptography.Aes.Create();
                aes.Key = AesKey;
                aes.IV  = AesIv;
                using var decryptor = aes.CreateDecryptor();
                var encryptedBytes = Convert.FromBase64String(encrypted);
                using var ms = new MemoryStream(encryptedBytes);
                using var cs = new System.Security.Cryptography.CryptoStream(ms, decryptor, System.Security.Cryptography.CryptoStreamMode.Read);
                using var reader = new StreamReader(cs);
                return reader.ReadToEnd();
            }
            catch { return null; }
        }

        private void SaveApiKey(string apiKey)
        {
            try
            {
                using var aes = System.Security.Cryptography.Aes.Create();
                aes.Key = AesKey;
                aes.IV  = AesIv;
                using var encryptor = aes.CreateEncryptor();
                using var ms = new MemoryStream();
                using var cs = new System.Security.Cryptography.CryptoStream(ms, encryptor, System.Security.Cryptography.CryptoStreamMode.Write);
                using var writer = new StreamWriter(cs);
                writer.Write(apiKey);
                writer.Flush();
                cs.FlushFinalBlock();
                var encrypted = Convert.ToBase64String(ms.ToArray());

                var obj = new JObject { ["claude_api_key"] = encrypted };
                File.WriteAllText(ApiSettingsFilePath, obj.ToString());
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
