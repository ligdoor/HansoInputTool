using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Newtonsoft.Json.Linq;
using NLog;

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

            // ② フォルダ選択ダイアログ（Windows IFileOpenDialog COM経由）
            var selectedBase = ShowFolderBrowserDialog("「HansoData」フォルダがある場所（親フォルダ）を選択してください");
            if (selectedBase == null)
            {
                MessageBox.Show("フォルダが選択されませんでした。\nアプリを終了します。",
                    "セットアップキャンセル", MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
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

            // vehicle_settings.json は空の状態で作成（既存は保持）
            var vehicleDest = Path.Combine(hansoDataPath, "vehicle_settings.json");
            if (!File.Exists(vehicleDest))
            {
                File.WriteAllText(vehicleDest, "{}");
                Logger.Info("vehicle_settings.json を作成しました（空）");
            }
        }

        // ─────────────────────────────────────────────────
        #region フォルダ選択ダイアログ（COMインターフェース経由）

        /// <summary>
        /// Windows IFileOpenDialog を使ったフォルダ選択ダイアログ。
        /// OpenFileDialogと違い「開く」ボタンで中に入らず正しくフォルダを選択できる。
        /// </summary>
        private static string ShowFolderBrowserDialog(string title)
        {
            var dialog = (IFileOpenDialog)new FileOpenDialog();
            try
            {
                dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                dialog.SetTitle(title);
                int hr = dialog.Show(IntPtr.Zero);
                if (hr < 0) return null; // キャンセル
                dialog.GetResult(out IShellItem item);
                item.GetDisplayName(SIGDN_FILESYSPATH, out string path);
                return path;
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }

        // COM定数・インターフェース定義
        private const uint FOS_PICKFOLDERS    = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint SIGDN_FILESYSPATH  = 0x80058000;

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialog { }

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            [PreserveSig] int Show(IntPtr parent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppsai);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        #endregion

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

        public static string ReadEraNameFromSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return "R";
                var json = File.ReadAllText(SettingsPath);
                var lines = json.Split('\n');
                var noComment = string.Join("\n",
                    System.Linq.Enumerable.Where(lines, l => !l.TrimStart().StartsWith("//")));
                var obj = JObject.Parse(noComment.Length > 2 ? noComment : "{}");
                var era = obj["EraName"]?.ToString();
                return string.IsNullOrWhiteSpace(era) ? "R" : era;
            }
            catch { return "R"; }
        }

        public static void SaveEraNameToSettings(string eraName)
        {
            try
            {
                JObject obj;
                if (File.Exists(SettingsPath))
                {
                    var existing = File.ReadAllText(SettingsPath);
                    var lines = existing.Split('\n');
                    var noComment = string.Join("\n",
                        System.Linq.Enumerable.Where(lines, l => !l.TrimStart().StartsWith("//")));
                    obj = JObject.Parse(noComment.Length > 2 ? noComment : "{}");
                }
                else
                {
                    obj = new JObject();
                }
                obj["EraName"] = eraName;
                File.WriteAllText(SettingsPath, obj.ToString(Newtonsoft.Json.Formatting.Indented));
                Logger.Info($"appsettings.json に EraName を保存: {eraName}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "appsettings.json の EraName 保存に失敗しました");
            }
        }

        public static int ReadEraStartYearFromSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return 2019; // 令和元年デフォルト
                var json = File.ReadAllText(SettingsPath);
                var lines = json.Split('\n');
                var noComment = string.Join("\n",
                    System.Linq.Enumerable.Where(lines, l => !l.TrimStart().StartsWith("//")));
                var obj = JObject.Parse(noComment.Length > 2 ? noComment : "{}");
                var val = obj["EraStartYear"]?.ToString();
                return int.TryParse(val, out int y) ? y : 2019;
            }
            catch { return 2019; }
        }

        public static void SaveEraStartYearToSettings(int eraStartYear)
        {
            try
            {
                JObject obj;
                if (File.Exists(SettingsPath))
                {
                    var existing = File.ReadAllText(SettingsPath);
                    var lines = existing.Split('\n');
                    var noComment = string.Join("\n",
                        System.Linq.Enumerable.Where(lines, l => !l.TrimStart().StartsWith("//")));
                    obj = JObject.Parse(noComment.Length > 2 ? noComment : "{}");
                }
                else
                {
                    obj = new JObject();
                }
                obj["EraStartYear"] = eraStartYear;
                File.WriteAllText(SettingsPath, obj.ToString(Newtonsoft.Json.Formatting.Indented));
                Logger.Info($"appsettings.json に EraStartYear を保存: {eraStartYear}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "appsettings.json の EraStartYear 保存に失敗しました");
            }
        }

        private static string ColumnMapPath
        {
            get
            {
                var dataPath = ReadDataPathFromSettings();
                return string.IsNullOrEmpty(dataPath)
                    ? null
                    : Path.Combine(dataPath, "column_map.json");
            }
        }

        public static HansoInputTool.Models.ColumnMapping ReadColumnMap()
        {
            try
            {
                var path = ColumnMapPath;
                if (path != null && File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var cm = Newtonsoft.Json.JsonConvert.DeserializeObject<HansoInputTool.Models.ColumnMapping>(json);
                    if (cm != null) return cm;
                }
            }
            catch (Exception ex) { Logger.Error(ex, "column_map.json の読み込みに失敗しました"); }

            // デフォルト値
            return new HansoInputTool.Models.ColumnMapping
            {
                NormalSheet = new HansoInputTool.Models.SheetColumnMap
                {
                    Day = 2, HansoCount = 3, YuryoKm = 4, MuryoKm = 5,
                    KihonFee = 6, SokoFee = 7, ShinyaFee = 8, TotalFee = 9, ShinyaMinutes = 11
                },
                EastSheet = new HansoInputTool.Models.CellAddressMap
                {
                    Jitsudo = "E4", Hanso = "G4", YuryoKm = "H4", MuryoKm = "I4", UnsoJisseki = "K4"
                },
                ShukeiSheet = new HansoInputTool.Models.CellAddressMap
                {
                    Days = "E4", Hanso = "G4", YuryoKm = "H4", MuryoKm = "I4", Total = "K4"
                }
            };
        }

        public static void SaveColumnMap(HansoInputTool.Models.ColumnMapping cm)
        {
            try
            {
                var path = ColumnMapPath;
                if (path == null) return;
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(cm, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(path, json);
                Logger.Info("column_map.json を保存しました");
            }
            catch (Exception ex) { Logger.Error(ex, "column_map.json の保存に失敗しました"); }
        }

        #region 期・R（前回値）の読み書き

        private static JObject LoadSettingsJson()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return new JObject();
                var json = File.ReadAllText(SettingsPath);
                var lines = json.Split('\n');
                var noComment = string.Join("\n",
                    System.Linq.Enumerable.Where(lines, l => !l.TrimStart().StartsWith("//")));
                return JObject.Parse(noComment.Length > 2 ? noComment : "{}");
            }
            catch { return new JObject(); }
        }

        public static (string period, string rNumber) ReadLastPeriodRNumber()
        {
            try
            {
                var obj = LoadSettingsJson();
                return (obj["LastPeriod"]?.ToString() ?? "", obj["LastRNumber"]?.ToString() ?? "");
            }
            catch { return ("", ""); }
        }

        public static void SaveLastPeriodRNumber(string period, string rNumber)
        {
            try
            {
                var obj = LoadSettingsJson();
                obj["LastPeriod"]  = period  ?? "";
                obj["LastRNumber"] = rNumber ?? "";
                File.WriteAllText(SettingsPath, obj.ToString(Newtonsoft.Json.Formatting.Indented));
                Logger.Info($"appsettings.json に LastPeriod={period}, LastRNumber={rNumber} を保存");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "LastPeriod/LastRNumber の保存に失敗しました");
            }
        }

        #endregion

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
