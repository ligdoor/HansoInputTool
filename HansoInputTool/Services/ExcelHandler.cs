using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// Excelファイルの読み書き・シート管理を担当するクラス。
    /// partial classで3ファイルに分割:
    ///   ExcelHandler.cs             - 基本操作・データ読み書き（このファイル）
    ///   ExcelHandler.SheetSync.cs   - シート同期・並べ替え・命名
    ///   ExcelHandler.MonthlySummary.cs - 月間集計シート更新
    /// </summary>
    public partial class ExcelHandler : IDisposable
    {
        private bool _disposed = false;

        static ExcelHandler()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly string _inputFilePath;
        private readonly string _templateFilePath;
        private ColumnMapping _columnMap;
        private ExcelPackage _inputPackage;
        private ExcelPackage _templatePackage;
        private readonly Dictionary<string, List<RowData>> _dataCache = new();

        // 動的フラグ管理サービス（外部から注入）
        public FlagDefinitionService FlagService { get; set; }

        // SQLiteデータベースサービス（外部から注入。nullのときはExcel読み書きにフォールバック）
        public DatabaseService DbService { get; set; }
        public VehicleSettingsService VehicleSettingsService { get; set; }

        public bool IsFeeMode(string sheetName)
            => VehicleSettingsService?.IsFeeMode(sheetName) ?? sheetName.Contains("大月");

        public void UpdateColumnMap(ColumnMapping newMap)
        {
            _columnMap = newMap;
        }

        public ExcelHandler(string inputFilePath, string templateFilePath, ColumnMapping columnMap)
        {
            _inputFilePath    = inputFilePath;
            _templateFilePath = templateFilePath;
            _columnMap        = columnMap;
            Load();
        }

        #region 基本操作

        public void Load()
        {
            _inputPackage?.Dispose();
            _templatePackage?.Dispose();
            _inputPackage    = new ExcelPackage(new FileInfo(_inputFilePath));
            _templatePackage = new ExcelPackage(new FileInfo(_templateFilePath));
            _dataCache.Clear();
        }

        /// <summary>
        /// 起動時にcustom_flags.jsonとInput.xlsxのヘッダー行を比較し、
        /// 差分があれば自動でフラグ列を同期する。
        /// バージョンアップ後の初回起動時に新旧フラグ構造を自動修復する。
        /// </summary>
        public void SyncFlagsOnStartup(FlagDefinitionService flagService)
        {
            if (flagService == null) return;
            var expectedFlags = flagService.Flags;
            if (expectedFlags.Count == 0) return;

            // 代表シート（最初の通常系シート）のヘッダー行を読んで現状を把握
            var targetSheets = _inputPackage.Workbook.Worksheets
                .Where(ws => (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車")
                           || ws.Name.Contains("CH") || IsTemplateSheet(ws.Name))
                          && !ws.Name.Contains("登録")
                          && ws.Name != "月間集計")
                .ToList();

            if (targetSheets.Count == 0) return;

            bool needsSave = false;
            foreach (var ws in targetSheets)
            {
                var addedFlags   = new List<FlagDefinition>();
                var removedFlags = new List<FlagDefinition>();

                // ヘッダー行（2行目）から現在のフラグ列を読む
                int lastCol = ws.Dimension?.End.Column ?? 0;
                var headerValues = new Dictionary<int, string>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var v = ws.Cells[2, c].Value?.ToString();
                    if (!string.IsNullOrEmpty(v)) headerValues[c] = v;
                }

                // expectedFlags にあるがヘッダーにない → 追加が必要
                foreach (var flag in expectedFlags)
                {
                    if (!headerValues.Values.Contains(flag.DisplayName))
                        addedFlags.Add(flag);
                }

                // [No.7修正] ヘッダーにある列名がexpectedFlagsのDisplayNameに存在しない → 削除対象。
                // 旧実装はfirstFlagCol以降という列番号ベースの判定だったため、
                // フラグの列番号が変わると無関係な列（日付・搬送回数等）を誤削除するリスクがあった。
                // 修正後はDisplayName完全一致で「フラグとして登録されている列名かどうか」だけを判断し、
                // さらにその列名が現在のexpectedFlagsにない場合のみ削除対象とする。
                var expectedDisplayNames = new HashSet<string>(expectedFlags.Select(f => f.DisplayName));
                var allFlagDisplayNames  = new HashSet<string>(expectedFlags.Select(f => f.DisplayName)); // 将来拡張用に分離
                foreach (var kv in headerValues)
                {
                    // expectedFlagsのいずれかのDisplayNameと一致する列名が
                    // 現在のexpectedFlagsには存在しない → 旧フラグ列として削除
                    bool wasEverAFlag = allFlagDisplayNames.Contains(kv.Value) == false
                        && expectedFlags.Any(f => f.DisplayName == kv.Value) == false;

                    // シンプルな判定: ヘッダー値がexpectedFlagsのDisplayNameリストに含まれない
                    // かつ、そのヘッダー値が以前フラグ列として存在していた痕跡（既存フラグIDと不一致）なら削除
                    // ここでは「expectedFlagsに存在しないDisplayNameを持つ列を全て削除対象」とする安全な実装に変更
                    if (!expectedDisplayNames.Contains(kv.Value))
                    {
                        // ただし固定ヘッダー列（日・搬送回数・有料キロ等）は削除しない。
                        // 固定列はexpectedFlagsのExcelColumnの最小値より左側にあると仮定し、
                        // かつヘッダー名がフラグっぽい（expectedFlagsのDisplayNameに近い）ものだけを削除対象とする。
                        // より安全のため: expectedFlagsが空でなく、かつその列番号がfirstFlagCol以上の場合のみ削除
                        int firstFlagCol = expectedFlags.Count > 0 ? expectedFlags.Min(f => f.ExcelColumn) : int.MaxValue;
                        if (kv.Key >= firstFlagCol)
                        {
                            removedFlags.Add(new FlagDefinition
                            {
                                DisplayName = kv.Value,
                                ExcelColumn = kv.Key
                            });
                        }
                    }
                }

                if (addedFlags.Count == 0 && removedFlags.Count == 0) continue;

                Logger.Info($"[{ws.Name}] 起動時フラグ同期: 追加={addedFlags.Count}件, 削除={removedFlags.Count}件");

                // 削除（降順）
                foreach (var flag in removedFlags.OrderByDescending(f => f.ExcelColumn))
                {
                    int col = flag.ExcelColumn;
                    if (col >= 1 && col <= (ws.Dimension?.End.Column ?? 0))
                    {
                        ws.DeleteColumn(col);
                        Logger.Info($"[{ws.Name}] 列{col}（{flag.DisplayName}）を削除");
                    }
                }

                // 追加（昇順）
                foreach (var flag in addedFlags.OrderBy(f => f.ExcelColumn))
                {
                    int col    = flag.ExcelColumn;
                    int srcCol = col - 1;
                    ws.InsertColumn(col, 1);
                    int lastRow = ws.Dimension?.End.Row ?? 50;
                    for (int row = 1; row <= lastRow; row++)
                        ws.Cells[row, col].StyleID = ws.Cells[row, srcCol].StyleID;
                    ws.Column(col).Width    = ws.Column(srcCol).Width;
                    ws.Cells[2, col].Value  = flag.DisplayName;
                    int dataStart = 3;
                    for (int row = dataStart; row <= lastRow; row++)
                        ws.Cells[row, col].Value = null;
                    Logger.Info($"[{ws.Name}] 列{col}（{flag.DisplayName}）を追加");
                }

                needsSave = true;
            }

            if (needsSave)
            {
                _inputPackage.Save();
                Logger.Info("起動時フラグ同期: Input.xlsxを保存しました。");
            }
        }

        public void Save()
        {
            _inputPackage?.Save();
            _templatePackage?.Save();
        }

        public bool TemplateSheetExists(string sheetName)
            => _templatePackage.Workbook.Worksheets.Any(ws => ws.Name == sheetName);

        public List<string> SheetNames => _inputPackage?.Workbook.Worksheets
            .Where(ws => !ws.Name.Contains("登録") && !ws.Name.Contains("給油管理") && !IsTemplateSheet(ws.Name))
            .Select(ws => ws.Name)
            .ToList() ?? new List<string>();

        /// <summary>
        /// Input.xlsx内にシートが存在するか確認（テンプレートシートも含む）
        /// </summary>
        public bool InputSheetExists(string sheetName)
            => _inputPackage?.Workbook.Worksheets.Any(ws => ws.Name == sheetName) ?? false;

        public List<string> GetVehicleSheetNames()
            => _inputPackage.Workbook.Worksheets
                .Where(s => !s.Name.Contains("登録") && !s.Name.Contains("給油管理") && !IsTemplateSheet(s.Name))
                .Select(ws =>
                {
                    var numStr = System.Text.RegularExpressions.Regex.Match(ws.Name, @"\d+$").Value;
                    int num    = int.TryParse(numStr, out var n) ? n : 0;
                    return new { Name = ws.Name, CategoryOrder = GetCategoryOrder(ws.Name), Num = num };
                })
                .OrderBy(v => v.CategoryOrder)
                .ThenBy(v => v.Num)
                .Select(v => v.Name)
                .ToList();

        /// <summary>
        /// 車両シートの並び順を変更してInput.xlsxのシート順も同期する。
        /// fromName のシートを toName の前後に移動する。
        /// </summary>
        public void MoveVehicleSheet(string fromName, string toName, bool insertBefore)
        {
            if (fromName == toName) return;
            try
            {
                if (insertBefore)
                    _inputPackage.Workbook.Worksheets.MoveBefore(fromName, toName);
                else
                    _inputPackage.Workbook.Worksheets.MoveAfter(fromName, toName);
                _inputPackage.Save();
                Logger.Info($"シート順変更: {fromName} → {(insertBefore ? "前" : "後")} of {toName}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"シート順変更失敗: {fromName}");
            }
        }

        #endregion

        #region データ読み取り（プレビュー用）

        public List<RowData> GetSheetDataForPreview(string sheetName)
        {
            if (sheetName == null) return new List<RowData>();

            // DBが注入済みの場合はDBから読み取る
            // ただし東日本シートはDBに保存されないためExcelから読む
            if (DbService != null && !sheetName.Contains("東日本"))
            {
                if (_dataCache.ContainsKey(sheetName)) return _dataCache[sheetName];
                var flags  = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();
                var result = DbService.GetSheetData(sheetName, flags);
                _dataCache[sheetName] = result;
                return result;
            }

            // DB未注入時はExcelから読み取る（フォールバック）
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                return new List<RowData>();
            if (_dataCache.ContainsKey(sheetName)) return _dataCache[sheetName];

            var ws            = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return new List<RowData>();

            var data       = new List<RowData>();
            var map        = _columnMap.NormalSheet;
            bool isOotsuki = IsFeeMode(sheetName);
            var flagDefs   = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();

            for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
            {
                if (ws.Cells[rowIndex, map.Day].Value == null && ws.Cells[rowIndex, map.YuryoKm].Value == null) continue;

                var flagValues = new Dictionary<string, int?>();
                foreach (var flag in flagDefs)
                    flagValues[flag.Id] = GetNullableInt(ws.Cells[rowIndex, flag.ExcelColumn].Value);

                var rowData = new RowData
                {
                    RowIndex         = rowIndex,
                    B_Day            = GetNullableInt(ws.Cells[rowIndex, map.Day].Value),
                    C_Hanso          = GetNullableInt(ws.Cells[rowIndex, map.HansoCount].Value),
                    D_YuryoKm        = GetNullableInt(ws.Cells[rowIndex, map.YuryoKm].Value),
                    E_MuryoKm        = GetNullableInt(ws.Cells[rowIndex, map.MuryoKm].Value),
                    H_LateFeeOotsuki = GetNullableInt(ws.Cells[rowIndex, map.ShinyaFee].Value),
                    K_LateMinutes    = GetNullableInt(ws.Cells[rowIndex, map.ShinyaMinutes].Value),
                    FlagValues       = flagValues,
                    FlagDefinitions  = flagDefs
                };
                rowData.LateValueText = isOotsuki
                    ? rowData.H_LateFeeOotsuki?.ToString()
                    : rowData.K_LateMinutes?.ToString();
                data.Add(rowData);
            }

            _dataCache[sheetName] = data;
            return data;
        }

        #endregion

        #region データ書き込み

        public (int targetRow, string insertInfo) RegisterNormalData(
            string sheetName,
            Dictionary<string, double?> values,
            Dictionary<string, bool> flagStates)
        {
            // DBが注入済みの場合はDBに書き込む
            if (DbService != null)
            {
                long dbId = DbService.InsertRecord(sheetName, values, flagStates, IsFeeMode(sheetName));
                InvalidateCache(sheetName);
                // rowIndexの代わりにdbIdを返す（呼び出し側はinsertInfoを表示するだけなので互換あり）
                return ((int)dbId, "");
            }

            // DB未注入時はExcelに書き込む（フォールバック）
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");

            var ws            = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) throw new Exception($"シート '{sheetName}' に '合計' 行が見つかりません。");

            var map = _columnMap.NormalSheet;
            int targetRow = -1;
            for (int r = 3; r < totalRowIndex; r++)
                if (ws.Cells[r, map.Day].Value == null) { targetRow = r; break; }

            string insertInfo = "";
            if (targetRow == -1)
            {
                ws.InsertRow(totalRowIndex, 1);
                targetRow  = totalRowIndex;
                insertInfo = "空き行がないため、合計行の上に新しい行を挿入します。";
            }

            WriteNormalValues(ws, targetRow, map, values, flagStates, IsFeeMode(sheetName));
            InvalidateCache(sheetName);
            return (targetRow, insertInfo);
        }

        public void UpdateNormalData(
            string sheetName,
            int rowIndex,
            Dictionary<string, double?> values,
            Dictionary<string, bool> flagStates)
        {
            // DBが注入済みの場合はDBを更新（rowIndexをdbIdとして使用）
            if (DbService != null)
            {
                DbService.UpdateRecord((long)rowIndex, sheetName, values, flagStates, IsFeeMode(sheetName));
                InvalidateCache(sheetName);
                return;
            }

            // DB未注入時はExcelを更新（フォールバック）
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            WriteNormalValues(
                _inputPackage.Workbook.Worksheets[sheetName],
                rowIndex, _columnMap.NormalSheet, values, flagStates,
                IsFeeMode(sheetName));
            InvalidateCache(sheetName);
        }

        /// <summary>
        /// 給油記録を1件登録する（給油管理対象車両のみ想定・DBモード専用）。
        /// transportRecordIdを指定すると、その搬送データ行に紐付けて登録する
        /// （同じ日に複数の搬送行があっても正しい行にだけ表示されるようにするため）。
        /// 転記時にInput.xlsxの給油管理表シートへ書き込まれる。
        /// </summary>
        public void RegisterFuelData(string sheetName, int day, double odometerKm, double liters, long? transportRecordId = null)
        {
            if (DbService == null)
                throw new InvalidOperationException("給油記録の登録にはデータベースモードが必要です。");

            DbService.InsertFuelRecord(sheetName, day, odometerKm, liters, transportRecordId);
            InvalidateCache(sheetName);
        }

        public void RegisterEastData(string sheetName, Dictionary<string, double?> values)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws  = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.EastSheet;
            ws.Cells[map.Jitsudo].Value     = values.GetValueOrDefault("延実働車輌数");
            ws.Cells[map.Hanso].Value       = values.GetValueOrDefault("搬送回数");
            ws.Cells[map.YuryoKm].Value     = values.GetValueOrDefault("有料キロ数");
            ws.Cells[map.MuryoKm].Value     = values.GetValueOrDefault("無料キロ数");
            ws.Cells[map.UnsoJisseki].Value = values.GetValueOrDefault("運輸実績");
        }

        public void DeleteRows(string sheetName, List<int> rowIndices)
        {
            // DBが注入済みの場合はDBから削除（rowIndexをdbIdとして使用）
            if (DbService != null)
            {
                foreach (var dbId in rowIndices)
                    DbService.DeleteRecord((long)dbId);
                InvalidateCache(sheetName);
                return;
            }

            // DB未注入時はExcelから削除（フォールバック）
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                throw new ArgumentException($"シートが見つかりません: {sheetName}");
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            foreach (var rowIndex in rowIndices.OrderByDescending(r => r))
                ws.DeleteRow(rowIndex);
            InvalidateCache(sheetName);
        }

        public List<string> ClearData()
        {
            var logMessages = new List<string>();
            var normalMap   = _columnMap.NormalSheet;
            var eastMap     = _columnMap.EastSheet;
            var flags       = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();

            // 固定列＋動的フラグ列の一覧を作成（基本料・走行料・合計も含める）
            var fixedCols  = new[] {
                normalMap.Day, normalMap.HansoCount,
                normalMap.YuryoKm, normalMap.MuryoKm,
                normalMap.KihonFee, normalMap.SokoFee,
                normalMap.ShinyaFee, normalMap.TotalFee,
                normalMap.ShinyaMinutes
            };
            var flagCols   = flags.Select(f => f.ExcelColumn).ToArray();
            var clearCols  = fixedCols.Concat(flagCols).Where(c => c > 0).Distinct().ToArray();

            foreach (var ws in _inputPackage.Workbook.Worksheets)
            {
                if (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                {
                    var totalRowIndex = FindTotalRow(ws);
                    if (totalRowIndex != -1)
                    {
                        // データ行（3行目〜合計行の直前）をクリア
                        for (int rowIndex = 3; rowIndex < totalRowIndex; rowIndex++)
                            foreach (int col in clearCols)
                                ws.Cells[rowIndex, col].Value = null;

                        // 合計行（103行目相当）もクリア
                        foreach (int col in clearCols)
                            ws.Cells[totalRowIndex, col].Value = null;

                        logMessages.Add($"[{ws.Name}] の入力値をクリアしました。");
                    }
                }
                else if (ws.Name.Contains("東日本"))
                {
                    ws.Cells[eastMap.Jitsudo].Value     = null;
                    ws.Cells[eastMap.Hanso].Value       = null;
                    ws.Cells[eastMap.YuryoKm].Value     = null;
                    ws.Cells[eastMap.MuryoKm].Value     = null;
                    ws.Cells[eastMap.UnsoJisseki].Value = null;
                    logMessages.Add($"[{ws.Name}] のデータをクリアしました。");
                }
            }
            _dataCache.Clear();
            return logMessages;
        }

        /// <summary>
        /// 東日本シートの登録済み値をセルアドレスから読み取って返す。
        /// 未入力の場合は null を返す。
        /// </summary>
        public Dictionary<string, double?> GetEastSheetValues(string sheetName)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName))
                return null;

            var ws  = _inputPackage.Workbook.Worksheets[sheetName];
            var map = _columnMap.EastSheet;

            return new Dictionary<string, double?>
            {
                ["延実働車輌数"] = GetNullableDouble(ws.Cells[map.Jitsudo].Value),
                ["搬送回数"]   = GetNullableDouble(ws.Cells[map.Hanso].Value),
                ["有料キロ数"] = GetNullableDouble(ws.Cells[map.YuryoKm].Value),
                ["無料キロ数"] = GetNullableDouble(ws.Cells[map.MuryoKm].Value),
                ["運輸実績"]   = GetNullableDouble(ws.Cells[map.UnsoJisseki].Value),
            };
        }

        /// <summary>
        /// フラグ定義の変更をInput.xlsxの全通常シートに反映する。
        /// 追加されたフラグ → 対象列を追加してヘッダーを書く
        /// 削除されたフラグ → 対象列を削除する
        /// 順番変更は列には反映しない
        /// </summary>
        public void SyncFlagColumns(
            IReadOnlyList<FlagDefinition> oldFlags,
            IReadOnlyList<FlagDefinition> newFlags)
        {
            // 追加されたフラグ（新規IDのもの）
            var addedFlags = newFlags
                .Where(n => !oldFlags.Any(o => o.Id == n.Id))
                .OrderBy(f => f.ExcelColumn)
                .ToList();

            // 削除されたフラグ（旧IDで新フラグに存在しないもの）
            var removedFlags = oldFlags
                .Where(o => !newFlags.Any(n => n.Id == o.Id))
                .OrderByDescending(f => f.ExcelColumn) // 後ろから削除して列番号ずれを防ぐ
                .ToList();

            if (addedFlags.Count == 0 && removedFlags.Count == 0) return;

            // ── Input.xlsx 対象シート ──────────────────────────────────────
            // 通常系（寝台車・霊柩車・CH）シート ＋ Input.xlsx末尾のTemplate1（ひな形シート）
            // 登録シート・月間集計は除外。Template1はひな形なので必ず含める。
            var allTargetSheets = _inputPackage.Workbook.Worksheets
                .Where(ws => !ws.Name.Contains("登録")
                          && ws.Name != "月間集計"
                          && !ws.Name.Contains("給油管理")  // 給油管理表はレイアウトが異なるため対象外
                          && (ws.Name.Contains("寝台車")
                           || ws.Name.Contains("霊柩車")
                           || ws.Name.Contains("CH")
                           || IsTemplateSheet(ws.Name)))  // Template1 を含む
                .ToList();

            Logger.Info($"SyncFlagColumns: 追加={addedFlags.Count}件, 削除={removedFlags.Count}件, " +
                        $"対象シート={allTargetSheets.Count}枚");

            foreach (var ws in allTargetSheets)
            {
                // --- 追加処理 ---
                foreach (var flag in addedFlags)
                {
                    int col = flag.ExcelColumn;

                    // 1列挿入
                    ws.InsertColumn(col, 1);

                    // 左隣の列（col-1）から書式・罫線をコピー
                    int srcCol = col - 1;
                    int totalRow = FindTotalRow(ws);
                    int lastRow = totalRow > 0 ? totalRow : ws.Dimension?.End.Row ?? 50;

                    for (int row = 1; row <= lastRow; row++)
                    {
                        var srcCell  = ws.Cells[row, srcCol];
                        var destCell = ws.Cells[row, col];

                        // 書式コピー
                        destCell.StyleID = srcCell.StyleID;
                    }

                    // 左隣と同じ列幅を設定（StyleIDでは列幅がコピーされないため個別に設定）
                    ws.Column(col).Width = ws.Column(srcCol).Width;

                    // 2行目（ヘッダー行）に表示名を記入
                    ws.Cells[2, col].Value = flag.DisplayName;

                    // データ行はクリア（書式だけ残す）
                    int dataStart = 3;
                    for (int row = dataStart; row <= lastRow; row++)
                        ws.Cells[row, col].Value = null;

                    Logger.Info($"[{ws.Name}] 列{col} にフラグ「{flag.DisplayName}」を追加しました。");
                }

                // --- 削除処理 ---
                foreach (var flag in removedFlags)
                {
                    int col = flag.ExcelColumn;
                    if (col < 1 || col > (ws.Dimension?.End.Column ?? 0)) continue;

                    ws.DeleteColumn(col);
                    Logger.Info($"[{ws.Name}] 列{col}（フラグ「{flag.DisplayName}」）を削除しました。");
                }
            }

            _dataCache.Clear();
        }

        public bool CheckRemainingData()
        {
            // DBが注入済みの場合はDBで確認
            if (DbService != null)
                return DbService.HasAnyData();

            // DB未注入時はExcelで確認（フォールバック）
            var map = _columnMap.NormalSheet;
            foreach (var ws in _inputPackage.Workbook.Worksheets)
                if ((ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                    && ws.Cells[3, map.Day].Value != null)
                    return true;
            return false;
        }

        #endregion

        #region 内部ヘルパー（他のpartialファイルからも使用）

        private void WriteNormalValues(ExcelWorksheet ws, int row, SheetColumnMap map,
            Dictionary<string, double?> values, Dictionary<string, bool> flagStates, bool isOotsuki)
        {
            double? yuryoVal = values.GetValueOrDefault("有料キロ(D)");
            int hansoVal     = (yuryoVal.HasValue && yuryoVal > 0) ? 1 : 0;

            ws.Cells[row, map.Day].Value        = values.GetValueOrDefault("日(B)");
            ws.Cells[row, map.HansoCount].Value = hansoVal;
            ws.Cells[row, map.YuryoKm].Value    = yuryoVal;
            ws.Cells[row, map.MuryoKm].Value    = values.GetValueOrDefault("無料キロ(E)");

            // 動的フラグを書き込む
            var flags = FlagService?.Flags ?? new List<Models.FlagDefinition>().AsReadOnly();
            foreach (var flag in flags)
            {
                bool isOn = flagStates != null && flagStates.TryGetValue(flag.Id, out bool v) && v;
                ws.Cells[row, flag.ExcelColumn].Value = isOn ? 1 : (object)null;
            }

            if (isOotsuki)
            {
                ws.Cells[row, map.ShinyaFee].Value     = values.GetValueOrDefault("深夜料金(H)");
                ws.Cells[row, map.ShinyaMinutes].Value = null;
            }
            else
            {
                ws.Cells[row, map.ShinyaFee].Value     = null;
                ws.Cells[row, map.ShinyaMinutes].Value = values.GetValueOrDefault("深夜時間(K)");
            }
        }

        public void InvalidateCache(string sheetName)
        {
            if (_dataCache.ContainsKey(sheetName)) _dataCache.Remove(sheetName);
        }

        /// <summary>全シートのキャッシュを破棄する（DBクリア後などに使用）</summary>
        public void InvalidateCacheAll()
        {
            _dataCache.Clear();
        }

        /// <summary>
        /// 読み込んだExcel（Input.xlsx）の通常系シートのデータをDBに一括インポートする。
        /// 「実績月報を読み込む」機能でExcelを差し替えた後に呼ぶ。
        /// DBの既存データはクリアしてからインポートする。
        /// </summary>
        public void ImportFromExcelToDb(DatabaseService dbService, FlagDefinitionService flagService)
        {
            if (dbService == null) return;

            var map   = _columnMap.NormalSheet;
            var flags = flagService?.Flags ?? new System.Collections.ObjectModel.ReadOnlyCollection<Models.FlagDefinition>(new System.Collections.Generic.List<Models.FlagDefinition>());

            // 通常系シートのみ対象（東日本・登録・テンプレート系は除外）
            var targetSheets = _inputPackage.Workbook.Worksheets
                .Where(ws => (ws.Name.Contains("寝台車") || ws.Name.Contains("霊柩車") || ws.Name.Contains("CH"))
                          && !ws.Name.Contains("登録")
                          && !IsTemplateSheet(ws.Name))
                .ToList();

            if (targetSheets.Count == 0) return;

            // 既存DBデータをクリア
            dbService.ClearAllData();
            Logger.Info($"ExcelからDBへインポート開始: {targetSheets.Count}シート");

            // 全シートのデータを収集してBulkInsertで一括登録（トランザクション）
            var allRecords = new System.Collections.Generic.List<(string, System.Collections.Generic.Dictionary<string, double?>, System.Collections.Generic.Dictionary<string, bool>, bool)>();

            foreach (var ws in targetSheets)
            {
                bool isOotsuki   = IsFeeMode(ws.Name);
                int  totalRowIdx = FindTotalRow(ws);
                if (totalRowIdx == -1) continue;

                for (int row = 3; row < totalRowIdx; row++)
                {
                    if (ws.Cells[row, map.Day].Value == null &&
                        ws.Cells[row, map.YuryoKm].Value == null) continue;

                    var values = new System.Collections.Generic.Dictionary<string, double?>
                    {
                        ["日(B)"]      = GetNullableDoubleFromCell(ws.Cells[row, map.Day].Value),
                        ["有料キロ(D)"] = GetNullableDoubleFromCell(ws.Cells[row, map.YuryoKm].Value),
                        ["無料キロ(E)"] = GetNullableDoubleFromCell(ws.Cells[row, map.MuryoKm].Value),
                        ["深夜料金(H)"] = isOotsuki ? GetNullableDoubleFromCell(ws.Cells[row, map.ShinyaFee].Value) : null,
                        ["深夜時間(K)"] = !isOotsuki ? GetNullableDoubleFromCell(ws.Cells[row, map.ShinyaMinutes].Value) : null,
                    };

                    var flagStates = new System.Collections.Generic.Dictionary<string, bool>();
                    foreach (var flag in flags)
                        flagStates[flag.Id] = GetNullableInt(ws.Cells[row, flag.ExcelColumn].Value) == 1;

                    allRecords.Add((ws.Name, values, flagStates, isOotsuki));
                }
            }

            // 1トランザクションで一括コミット（1件ずつより大幅に高速）
            dbService.BulkInsert(allRecords);
            _dataCache.Clear();
            Logger.Info($"ExcelからDBへインポート完了: {allRecords.Count}行");
        }

        private static double? GetNullableDoubleFromCell(object cellValue)
        {
            if (cellValue == null) return null;
            if (cellValue is double d) return d;
            if (double.TryParse(cellValue.ToString(), out var result)) return result;
            return null;
        }

        /// <summary>
        /// 指定シートの指定フラグ列がON(=1)の行数を返す
        /// </summary>
        public int GetFlagCount(string sheetName, int excelColumn)
        {
            if (!_inputPackage.Workbook.Worksheets.Any(s => s.Name == sheetName)) return 0;
            var ws = _inputPackage.Workbook.Worksheets[sheetName];
            var totalRowIndex = FindTotalRow(ws);
            if (totalRowIndex == -1) return 0;
            int count = 0;
            for (int r = 3; r < totalRowIndex; r++)
                if (GetNullableInt(ws.Cells[r, excelColumn].Value) == 1) count++;
            return count;
        }

        /// <summary>後方互換：エンバーミング件数</summary>
        public int GetEmbalmingCount(string sheetName)
        {
            var embFlag = FlagService?.Flags.FirstOrDefault(f => f.Id == "embalming");
            int col = embFlag?.ExcelColumn ?? _columnMap.NormalSheet.IsEmbalming;
            return GetFlagCount(sheetName, col);
        }

        internal static int FindTotalRow(ExcelWorksheet ws)
        {
            if (ws?.Dimension == null) return -1;
            for (int row = ws.Dimension.End.Row; row >= 3; row--)
                if (ws.Cells[row, 1].Value?.ToString()?.Contains("合計") == true) return row;
            return -1;
        }

        internal static int? GetNullableInt(object val)
        {
            if (val == null) return null;
            if (val is int i)     return i;
            if (val is long l)    return (int)l;
            if (val is double d)  return (int)d;
            if (val is decimal m) return (int)m;

            var s = val.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            s = s.Replace(",", "").Replace("，", "");

            if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double parsed) ||
                double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsed))
                return (int)parsed;

            Logger.Warn($"GetNullableInt: 非数値フィールドをパースできませんでした: '{s}'");
            return null;
        }

        internal static double? GetNullableDouble(object val)
            => val == null ? null : Convert.ToDouble(val);

        internal static bool IsTemplateSheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (sheetName.Replace(" ", "").StartsWith("template", StringComparison.OrdinalIgnoreCase)) return true;
            if (sheetName.IndexOf("テンプレート", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        internal static bool NeedQuotes(string sheetName)
            => sheetName.Contains(" ") || sheetName.Contains("-") || sheetName.Contains("(")
            || sheetName.Contains(")") || sheetName.Contains("'") || sheetName.Contains("!")
            || sheetName.Contains("#");

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                _inputPackage?.Dispose();
                _templatePackage?.Dispose();
                _disposed = true;
            }
        }
    }
}
