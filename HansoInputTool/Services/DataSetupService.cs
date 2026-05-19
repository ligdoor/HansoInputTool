using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using NLog;
using Shell = System.Runtime.InteropServices;

namespace HansoInputTool.Services
{
    /// <summary>
    /// 初回起動時またはappsettings.jsonが未設定の場合に
    /// HansoDataフォルダの場所をスタッフに選択させ、
    /// 必要なファイルを配置するセットアップサービス。
    /// </summary>
    public static class DataSetupService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static readonly string AppDir          = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string SettingsPath    = Path.Combine(AppDir, "appsettings.json");

        // HansoDataフォルダ内に必要な初期ファイル（exeと同じdata/フォルダからコピー）
        private static readonly string[] RequiredFiles =
        {
            "Input.xlsx",
            "Template.xlsx",
            "column_map.json",
            "rates.json",
            "shortcuts.json",
        };

        /// <summary>
        /// 起動時に呼び出す。
        /// appsettings.jsonにDataPathが設定済みかつフォルダが存在すれば何もしない。
        /// 未設定またはフォルダが存在しない場合はセットアップを実行する。
        /// </summary>
        /// <returns>セットアップが完了したデータパス。失敗/キャンセル時は null。</returns>
        public static string EnsureDataPath()
        {
            // 既存のappsettings.jsonを読む
            var existingPath = ReadDataPathFromSettings();
            if (!string.IsNullOrWhiteSpace(existingPath) && Directory.Exists(existingPath))
            {
                Logger.Info($"DataPath確認済み: {existingPath}");
                return existingPath;
            }

            // 未設定 or フォルダが消えている → セットアップ開始
            Logger.Info("DataPath未設定または存在しないため、セットアップを開始します。");
            return RunSetup(existingPath);
        }

        // ─────────────────────────────────────────────────
        private static string RunSetup(string previousPath)
        {
            // ① 案内メッセージ
            var msg = string.IsNullOrWhiteSpace(previousPath)
                ? "HansoInputToolへようこそ！\n\n" +
                  "データファイルの保存先フォルダを選択してください。\n" +
                  "選択したフォルダ内の「HansoData」フォルダを使用します。\n\n" +
                  "※ 複数のPCで共有する場合はサーバーの共有フォルダを選択してください。"
                : "前回設定されたデータフォルダが見つかりません。\n\n" +
                  $"前回のパス: {previousPath}\n\n" +
                  "HansoDataフォルダがある場所を選択してください。";

            MessageBox.Show(msg, "データフォルダの設定",
                MessageBoxButton.OK, MessageBoxImage.Information);

            // ② フォルダ選択ダイアログ（WPFネイティブ：OpenFileDialogをフォルダ選択に流用）
            var dialog = new OpenFileDialog
            {
                Title            = "「HansoData」フォルダがある場所（親フォルダ）を選択してください",
                Filter           = "フォルダを選択|*.none",
                FileName         = "フォルダを選択してください",
                CheckFileExists  = false,
                CheckPathExists  = true,
                ValidateNames    = false,
            };

            if (dialog.ShowDialog() != true)
            {
                MessageBox.Show("フォルダが選択されませんでした。\nアプリを終了します。",
                    "セットアップキャンセル", MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }

            // ファイル名部分を除いてフォルダパスを取得
            var selectedBase  = Path.GetDirectoryName(dialog.FileName) ?? dialog.FileName;
            var hansoDataPath = Path.Combine(selectedBase, "HansoData");

            // ③ 既存HansoDataフォルダが見つかった場合 → そのまま使用
            if (Directory.Exists(hansoDataPath))
            {
                var existingFiles = Directory.GetFiles(hansoDataPath);
                if (existingFiles.Length > 0)
                {
                    Logger.Info($"既存HansoDataフォルダを検出: {hansoDataPath} ({existingFiles.Length}ファイル)");

                    // appsettings.jsonにパスを書くだけ（ファイルコピーなし）
                    SaveDataPathToSettings(hansoDataPath);

                    MessageBox.Show(
                        $"既存のデータフォルダを使用します。\n\n{hansoDataPath}\n\nアプリを起動します。",
                        "設定完了", MessageBoxButton.OK, MessageBoxImage.Information);

                    Logger.Info("既存HansoDataフォルダを使用（ファイルコピーなし）");
                    return hansoDataPath;
                }
            }

            // ④ HansoDataフォルダが存在しない → 新規作成
            try
            {
                Directory.CreateDirectory(hansoDataPath);
                Logger.Info($"HansoDataフォルダ作成: {hansoDataPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"フォルダの作成に失敗しました。\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            // ⑤ 初期ファイルをコピー（新規作成時のみ）
            CopyInitialFiles(hansoDataPath);

            // ⑥ appsettings.json に保存
            SaveDataPathToSettings(hansoDataPath);

            MessageBox.Show(
                $"データフォルダを新規作成しました。\n\n{hansoDataPath}\n\nアプリを起動します。",
                "セットアップ完了", MessageBoxButton.OK, MessageBoxImage.Information);

            Logger.Info($"セットアップ完了（新規作成）: {hansoDataPath}");
            return hansoDataPath;
        }

        private static void CopyInitialFiles(string hansoDataPath)
        {
            // exeフォルダの "data" サブフォルダから初期ファイルをコピー
            var sourceDir = Path.Combine(AppDir, "data");

            foreach (var fileName in RequiredFiles)
            {
                var dest = Path.Combine(hansoDataPath, fileName);
                if (File.Exists(dest))
                {
                    Logger.Info($"スキップ（既存）: {fileName}");
                    continue;
                }

                var src = Path.Combine(sourceDir, fileName);
                if (File.Exists(src))
                {
                    File.Copy(src, dest);
                    Logger.Info($"コピー: {fileName}");
                }
                else
                {
                    Logger.Warn($"初期ファイルが見つかりません（スキップ）: {src}");
                }
            }

            // custom_flags.json は空の状態で作成（既存は保持）
            var flagsDest = Path.Combine(hansoDataPath, "custom_flags.json");
            if (!File.Exists(flagsDest))
            {
                File.WriteAllText(flagsDest, "[]");
                Logger.Info("custom_flags.json を作成しました（空）");
            }
        }

        // ─────────────────────────────────────────────────
        #region appsettings.json の読み書き

        public static string ReadDataPathFromSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return null;
                var json = File.ReadAllText(SettingsPath);
                var obj  = JObject.Parse(json);
                var raw  = obj["DataPath"]?.ToString();
                if (string.IsNullOrWhiteSpace(raw)) return null;

                // 相対パスは絶対パスに変換
                return Path.IsPathRooted(raw)
                    ? raw
                    : Path.GetFullPath(Path.Combine(AppDir, raw));
            }
            catch { return null; }
        }

        public static void SaveDataPathToSettings(string dataPath)
        {
            try
            {
                // 既存のappsettings.jsonがあれば読み込んで DataPath だけ更新
                JObject obj;
                if (File.Exists(SettingsPath))
                {
                    var existing = File.ReadAllText(SettingsPath);
                    // コメント行を除去してからパース
                    var lines    = existing.Split('\n');
                    var noComment = string.Join("\n",
                        System.Linq.Enumerable.Where(lines,
                            l => !l.TrimStart().StartsWith("//")));
                    obj = JObject.Parse(noComment.Length > 2 ? noComment : "{}");
                }
                else
                {
                    obj = new JObject();
                }

                obj["DataPath"] = dataPath;
                File.WriteAllText(SettingsPath, obj.ToString(Newtonsoft.Json.Formatting.Indented));
                Logger.Info($"appsettings.json に DataPath を保存: {dataPath}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "appsettings.json の保存に失敗しました");
            }
        }

        #endregion
    }
}
