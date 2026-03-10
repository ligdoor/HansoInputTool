using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NLog;

namespace HansoInputTool.Services
{
    public class BackupService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly string _backupDir;
        // バックアップ保持数（設定画面から変更可能）
        public int MaxBackupFiles { get; set; } = 10;
        public int MaxManualBackupFiles { get; set; } = 20;

        public BackupService()
        {
            _backupDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backups");
            Directory.CreateDirectory(_backupDir);
            Logger.Info($"バックアップディレクトリ: {_backupDir}");
        }

        /// <summary>
        /// ファイルのバックアップを作成
        /// </summary>
        public string CreateBackup(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Logger.Warn($"バックアップ対象のファイルが見つかりません: {filePath}");
                    return null;
                }

                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var extension = Path.GetExtension(filePath);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupFileName = $"{fileName}_{timestamp}{extension}";
                var backupPath = Path.Combine(_backupDir, backupFileName);

                File.Copy(filePath, backupPath, true);
                Logger.Info($"バックアップを作成しました: {backupFileName}");

                // 古いバックアップを削除
                CleanOldBackups(fileName, extension);

                return backupPath;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"バックアップの作成中にエラーが発生しました: {filePath}");
                return null;
            }
        }

        /// <summary>
        /// 自動バックアップ（アプリ起動時）
        /// </summary>
        public void CreateAutoBackup(string filePath)
        {
            var backupPath = CreateBackup(filePath);
            if (backupPath != null)
            {
                Logger.Info($"自動バックアップ完了: {Path.GetFileName(backupPath)}");
            }
        }

        /// <summary>
        /// 手動バックアップ（ユーザーが明示的に実行）
        /// </summary>
        public string CreateManualBackup(string filePath, string description = "")
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Logger.Warn($"バックアップ対象のファイルが見つかりません: {filePath}");
                    return null;
                }

                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var extension = Path.GetExtension(filePath);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var descSuffix = string.IsNullOrWhiteSpace(description) ? "" : $"_{description}";
                var backupFileName = $"{fileName}_{timestamp}{descSuffix}_manual{extension}";
                var backupPath = Path.Combine(_backupDir, backupFileName);

                File.Copy(filePath, backupPath, true);
                Logger.Info($"手動バックアップを作成しました: {backupFileName}");

                // 古い手動バックアップを削除
                CleanOldManualBackups(fileName, extension);

                return backupPath;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"手動バックアップの作成中にエラーが発生しました: {filePath}");
                return null;
            }
        }

        /// <summary>
        /// 古いバックアップを削除（最新N件のみ保持）
        /// </summary>
        private void CleanOldBackups(string baseFileName, string extension)
        {
            try
            {
                var pattern = $"{baseFileName}_*{extension}";
                var backups = Directory.GetFiles(_backupDir, pattern)
                    .Where(f => !f.Contains("_manual")) // 手動バックアップは除外
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .Skip(MaxBackupFiles)
                    .ToList();

                foreach (var oldBackup in backups)
                {
                    File.Delete(oldBackup);
                    Logger.Info($"古いバックアップを削除しました: {Path.GetFileName(oldBackup)}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "古いバックアップの削除中にエラーが発生しました");
            }
        }

        /// <summary>
        /// 古い手動バックアップを削除（最新N件のみ保持）
        /// </summary>
        private void CleanOldManualBackups(string baseFileName, string extension)
        {
            try
            {
                var pattern = $"{baseFileName}_*_manual{extension}";
                var backups = Directory.GetFiles(_backupDir, pattern)
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .Skip(MaxManualBackupFiles)
                    .ToList();

                foreach (var oldBackup in backups)
                {
                    File.Delete(oldBackup);
                    Logger.Info($"古い手動バックアップを削除しました: {Path.GetFileName(oldBackup)}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "古い手動バックアップの削除中にエラーが発生しました");
            }
        }

        /// <summary>
        /// 利用可能なバックアップ一覧を取得
        /// </summary>
        public List<BackupInfo> GetAvailableBackups(string baseFileName)
        {
            try
            {
                var pattern = $"{baseFileName}_*.xlsx";
                var backups = Directory.GetFiles(_backupDir, pattern)
                    .Select(f => new BackupInfo
                    {
                        FilePath = f,
                        FileName = Path.GetFileName(f),
                        CreatedTime = File.GetCreationTime(f),
                        FileSize = new FileInfo(f).Length,
                        IsManual = f.Contains("_manual")
                    })
                    .OrderByDescending(b => b.CreatedTime)
                    .ToList();

                return backups;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "バックアップ一覧の取得中にエラーが発生しました");
                return new List<BackupInfo>();
            }
        }

        /// <summary>
        /// バックアップから復元
        /// </summary>
        public bool RestoreFromBackup(string backupPath, string targetPath)
        {
            try
            {
                if (!File.Exists(backupPath))
                {
                    Logger.Warn($"復元元のバックアップファイルが見つかりません: {backupPath}");
                    return false;
                }

                // 復元前に現在のファイルをバックアップ
                if (File.Exists(targetPath))
                {
                    CreateBackup(targetPath);
                }

                File.Copy(backupPath, targetPath, true);
                Logger.Info($"バックアップから復元しました: {Path.GetFileName(backupPath)} → {Path.GetFileName(targetPath)}");

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"バックアップからの復元中にエラーが発生しました: {backupPath}");
                return false;
            }
        }

        /// <summary>
        /// バックアップフォルダを開く
        /// </summary>
        public void OpenBackupFolder()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = _backupDir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "バックアップフォルダを開く際にエラーが発生しました");
            }
        }
    }

    public class BackupInfo
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public DateTime CreatedTime { get; set; }
        public long FileSize { get; set; }
        public bool IsManual { get; set; }

        public string DisplayName => $"{FileName} ({FormatFileSize(FileSize)}) - {(IsManual ? "手動" : "自動")}";
        public string DisplayTime => CreatedTime.ToString("yyyy/MM/dd HH:mm:ss");

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }
}