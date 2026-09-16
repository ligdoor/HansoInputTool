using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HansoInputTool.Services
{
    public class MonthlyRecord
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string ShishaName { get; set; } = "";   // 支社名
        public string VehicleNo  { get; set; } = "";   // 車両番号
        public string VehicleKey   => $"{ShishaName}_{VehicleNo}";
        public string VehicleLabel => $"{ShishaName} {VehicleNo}";
        public double? JitsuzaiSuu { get; set; }        // D列：延実在車輌数
        public double? JitsudouSuu { get; set; }        // E列：延実働車輌数
        public double? Hanso       { get; set; }        // G列：搬送回数
        public double? YuryoKm     { get; set; }        // H列：有料キロ数
        public double? MuryoKm     { get; set; }        // I列：無料キロ数
        public double? Unshu       { get; set; }        // K列：運輸実績
    }

    /// <summary>
    /// チェックリスト表示用の車両エントリ
    /// </summary>
    public class VehicleEntry
    {
        public string Key       { get; set; } = "";   // "{ShishaName}_{VehicleNo}"
        public string Label     { get; set; } = "";   // 表示名
        public string ShishaName{ get; set; } = "";
        public string VehicleNo { get; set; } = "";
        public bool   IsKnown   { get; set; }         // SheetNameMap/FullSheetPatternで解決できた車両
        public bool   IsChecked { get; set; } = true; // チェック状態（デフォルトON）
    }

    public class VehicleAnnualSummaryService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private const string TargetSheetName = "月間集計";
        private const int DataStartRow = 4;
        private const int ColJitsuzai = 4;   // D: 延実在車輌数
        private const int ColJitsudou = 5;   // E: 延実働車輌数
        private const int ColHanso    = 7;   // G: 搬送回数
        private const int ColYuryoKm  = 8;   // H: 有料キロ数
        private const int ColMuryoKm  = 9;   // I: 無料キロ数
        private const int ColUnshu    = 11;  // K: 運輸実績

        private static readonly Regex FilePattern = new Regex(
            @"\d+期\s+(\d+)月\s+([A-Za-z]{1,3})(\d+)\s+アルス搬送・霊柩車\u3000実績月報集計\.xlsx$",
            RegexOptions.Compiled);

        // 短縮シート名 → (支社名, 車番) マッピング
        private static readonly Dictionary<string, (string Shisha, string VehicleNo)> SheetNameMap =
            new Dictionary<string, (string, string)>
            {
                ["寝台車 29"]        = ("CH富士吉田", "29"),
                ["寝台車 30"]        = ("CH富士吉田", "30"),
                ["霊柩車 40"]        = ("CH富士吉田", "40"),
                ["霊柩車 223"]       = ("CH富士吉田", "223"),
                ["大月 寝台車 1603"] = ("CH大月", "1603"),
                ["大月 霊柩車 2577"] = ("CH大月", "2577"),
                ["東日本セレモニー 2"] = ("東日本セレモニー", "2"),
            };

        private static readonly Regex FullSheetPattern = new Regex(
            @"^(CH富士吉田|CH大月|CH東富士|東日本セレモニー)(?:\s+(?:寝台車|霊柩車))?\s+(\d+)$");

        private static readonly Dictionary<string, int> CategoryOrder =
            new Dictionary<string, int>
            {
                ["CH富士吉田"]       = 1,
                ["CH大月"]           = 2,
                ["CH東富士"]         = 3,
                ["東日本セレモニー"] = 4,
            };

        // ─────────────────────────────────────
        // Step1: フォルダをスキャンして車両リストを収集する
        // （チェックリスト表示用。未知の車両もUnknownとして含める）
        // ─────────────────────────────────────
        public List<VehicleEntry> ScanVehicles(
            string rootFolder,
            int startYear, int startMonth,
            int endYear,   int endMonth)
        {
            var entries = new Dictionary<string, VehicleEntry>(); // key→entry

            if (!Directory.Exists(rootFolder)) return new List<VehicleEntry>();

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var files = Directory.GetFiles(rootFolder, "*実績月報集計*.xlsx",
                                           SearchOption.AllDirectories);

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var match = FilePattern.Match(fileName);
                if (!match.Success) continue;

                int month      = int.Parse(match.Groups[1].Value);
                string fileEra = match.Groups[2].Value;
                int eraNum     = int.Parse(match.Groups[3].Value);

                int baseYear = fileEra.Equals(eraName, StringComparison.OrdinalIgnoreCase)
                    ? eraStartYear - 1 : 2018;
                int year = baseYear + eraNum;

                if (!IsInRange(year, month, startYear, startMonth, endYear, endMonth))
                    continue;

                try
                {
                    // [EPPlus 8対応] LicenseContextはEPPlus 8で非推奨（obsolete）になったため、
                    // 新しいLicense APIに変更。※「アルス」の部分は実際の組織名に置き換えてください。
                    ExcelPackage.License.SetNonCommercialOrganization("アルス");
                    using var pkg = new ExcelPackage(new FileInfo(filePath));

                    foreach (var ws in pkg.Workbook.Worksheets)
                    {
                        var sname = ws.Name;
                        if (sname == TargetSheetName || sname.Contains("登録") ||
                            sname == "Template" || sname == "月間集計") continue;

                        bool isKnown = TryParseSheetName(sname, out var shisha, out var vehicleNo);

                        if (!isKnown)
                        {
                            // 未知シートも「未分類」として取り込む
                            shisha    = "未分類";
                            vehicleNo = sname;
                            Logger.Warn($"未知のシート名（未分類として追加）: [{sname}]");
                        }

                        var key = $"{shisha}_{vehicleNo}";
                        if (!entries.ContainsKey(key))
                        {
                            entries[key] = new VehicleEntry
                            {
                                Key        = key,
                                Label      = isKnown ? $"{shisha} {vehicleNo}" : $"[未分類] {sname}",
                                ShishaName = shisha,
                                VehicleNo  = vehicleNo,
                                IsKnown    = isKnown,
                                IsChecked  = true,
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"車両スキャンエラー: {filePath}");
                }
            }

            // ソート: 既知→未分類、支社順、車番順
            return entries.Values
                .OrderBy(e => e.IsKnown ? 0 : 1)
                .ThenBy(e => CategoryOrder.TryGetValue(e.ShishaName, out int c) ? c : 99)
                .ThenBy(e => int.TryParse(e.VehicleNo, out int n) ? n : 0)
                .ThenBy(e => e.VehicleNo)
                .ToList();
        }

        // ─────────────────────────────────────
        // Step2: チェックされた車両のみデータを読み込む
        // ─────────────────────────────────────
        public List<MonthlyRecord> LoadData(
            string rootFolder,
            int startYear, int startMonth,
            int endYear,   int endMonth,
            IEnumerable<VehicleEntry> selectedVehicles)
        {
            var result = new List<MonthlyRecord>();
            var selectedKeys = new HashSet<string>(selectedVehicles.Select(v => v.Key));
            if (selectedKeys.Count == 0) return result;

            if (!Directory.Exists(rootFolder)) return result;

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var files = Directory.GetFiles(rootFolder, "*実績月報集計*.xlsx",
                                           SearchOption.AllDirectories);

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var match = FilePattern.Match(fileName);
                if (!match.Success) continue;

                int month      = int.Parse(match.Groups[1].Value);
                string fileEra = match.Groups[2].Value;
                int eraNum     = int.Parse(match.Groups[3].Value);

                int baseYear = fileEra.Equals(eraName, StringComparison.OrdinalIgnoreCase)
                    ? eraStartYear - 1 : 2018;
                int year = baseYear + eraNum;

                if (!IsInRange(year, month, startYear, startMonth, endYear, endMonth))
                    continue;

                try
                {
                    // [EPPlus 8対応] LicenseContextはEPPlus 8で非推奨（obsolete）になったため、
                    // 新しいLicense APIに変更。※「アルス」の部分は実際の組織名に置き換えてください。
                    ExcelPackage.License.SetNonCommercialOrganization("アルス");
                    using var pkg = new ExcelPackage(new FileInfo(filePath));

                    foreach (var ws in pkg.Workbook.Worksheets)
                    {
                        var sname = ws.Name;
                        if (sname == TargetSheetName || sname.Contains("登録") ||
                            sname == "Template" || sname == "月間集計") continue;

                        bool isKnown = TryParseSheetName(sname, out var shisha, out var vehicleNo);
                        if (!isKnown) { shisha = "未分類"; vehicleNo = sname; }

                        var key = $"{shisha}_{vehicleNo}";
                        if (!selectedKeys.Contains(key)) continue; // チェックOFFはスキップ

                        var unshu = GetDouble(ws.Cells[DataStartRow, ColUnshu].Value);
                        var jitsuzai = GetDouble(ws.Cells[DataStartRow, ColJitsuzai].Value);
                        var jitsudou = GetDouble(ws.Cells[DataStartRow, ColJitsudou].Value);
                        var hanso    = GetDouble(ws.Cells[DataStartRow, ColHanso].Value);
                        var yuryoKm  = GetDouble(ws.Cells[DataStartRow, ColYuryoKm].Value);
                        var muryoKm  = GetDouble(ws.Cells[DataStartRow, ColMuryoKm].Value);
                        if (unshu == null && jitsuzai == null && jitsudou == null &&
                            hanso == null && yuryoKm == null && muryoKm == null) continue;

                        result.Add(new MonthlyRecord
                        {
                            Year        = year,
                            Month       = month,
                            ShishaName  = shisha,
                            VehicleNo   = vehicleNo,
                            JitsuzaiSuu = jitsuzai,
                            JitsudouSuu = jitsudou,
                            Hanso       = hanso,
                            YuryoKm     = yuryoKm,
                            MuryoKm     = muryoKm,
                            Unshu       = unshu,
                        });
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, $"データ読み込みエラー: {filePath}");
                }
            }

            return result;
        }

        // ─────────────────────────────────────
        // Step3: 選択車両でExcel出力
        // annual_results.xlsx（ひな形）を実際に開き、その中で「Template」シートを
        // 月ごとに複製して「月間集計 R#.#月」を作成、最後に「年間実績」シートへ
        // 全期間を3D集計する。デザイン（テーブル・配色・書式）はひな形のものをそのまま使う。
        // ─────────────────────────────────────
        public void ExportToExcel(
            List<MonthlyRecord> allData,
            List<VehicleEntry>  selectedVehicles,
            string outputPath,
            int startYear, int startMonth,
            int endYear,   int endMonth,
            string annualTemplatePath)
        {
            if (string.IsNullOrWhiteSpace(annualTemplatePath) || !File.Exists(annualTemplatePath))
                throw new FileNotFoundException($"年間集計のひな形ファイルが見つかりません: {annualTemplatePath}");

            // [EPPlus 8対応] LicenseContextはEPPlus 8で非推奨（obsolete）になったため、
            // 新しいLicense APIに変更。※「アルス」の部分は実際の組織名に置き換えてください。
            ExcelPackage.License.SetNonCommercialOrganization("アルス");

            // [デザイン維持対応] ひな形ファイル自体を開いて作業する。EPPlusには「テーブルを含む
            // シートを別ファイルへコピーするとファイルが壊れる」既知の不具合があるため、
            // 新規シートの複製はすべて「同じファイル（同じパッケージ）内」で行い、
            // 最後に別名（outputPath）で保存する（ひな形ファイル自体は一切上書きしない）。
            using var pkg = new ExcelPackage(new FileInfo(annualTemplatePath));

            var templateWs = pkg.Workbook.Worksheets["Template"];
            var summaryWs  = pkg.Workbook.Worksheets["年間実績"];
            if (templateWs == null || summaryWs == null)
                throw new InvalidOperationException(
                    "ひな形ファイル(annual_results.xlsx)に「Template」または「年間実績」シートが見つかりません。");

            string eraName      = DataSetupService.ReadEraNameFromSettings();
            int    eraStartYear = DataSetupService.ReadEraStartYearFromSettings();

            var months = GetMonthRange(startYear, startMonth, endYear, endMonth);
            int vehicleCount = selectedVehicles.Count;
            int totalDataRow = DataStartRow + vehicleCount; // 合計行

            var monthlySheetNames = new List<string>();

            // ---- 月ごとに「Template」シートを複製 ----
            foreach (var (y, m) in months)
            {
                int eraNum = y - (eraStartYear - 1);
                string sheetName = $"月間集計 {eraName}{eraNum}.{m}月";
                monthlySheetNames.Add(sheetName);

                var ws = pkg.Workbook.Worksheets.Add(sheetName, templateWs);
                ResizeSheetTable(ws, vehicleCount);

                ws.Cells[1, 1].Value = $"{eraName}{eraNum}";
                ws.Cells[1, 2].Value = m; // B1：月の数値
                ws.Cells[1, 3].Value = "月分";

                for (int vi = 0; vi < vehicleCount; vi++)
                {
                    var v = selectedVehicles[vi];
                    int row = DataStartRow + vi;
                    var record = allData.FirstOrDefault(d =>
                        d.VehicleKey == v.Key && d.Year == y && d.Month == m);

                    WriteVehicleRowValues(ws, row, vi + 1, v.ShishaName, v.VehicleNo, record);
                }

                WriteTotalRowFormulas(ws, totalDataRow, DataStartRow, totalDataRow - 1);
            }

            // ---- 年間実績シート（ひな形に元から入っているシートを流用） ----
            ResizeSheetTable(summaryWs, vehicleCount);

            int startEraNum = startYear - (eraStartYear - 1);
            int endEraNum   = endYear   - (eraStartYear - 1);
            summaryWs.Cells[1, 1].Value =
                $"{eraName}{startEraNum}年{startMonth}月～{eraName}{endEraNum}年{endMonth}月";

            string sheetRange3D = monthlySheetNames.Count == 1
                ? $"'{monthlySheetNames[0]}'"
                : $"'{monthlySheetNames[0]}:{monthlySheetNames[^1]}'";

            for (int vi = 0; vi < vehicleCount; vi++)
            {
                var v = selectedVehicles[vi];
                int row = DataStartRow + vi;

                summaryWs.Cells[row, 1].Value = $"№{vi + 1}";
                summaryWs.Cells[row, 2].Value = v.ShishaName;
                summaryWs.Cells[row, 3].Value = int.TryParse(v.VehicleNo, out int vn) ? (object)vn : v.VehicleNo;

                summaryWs.Cells[row, 4].Formula  = $"SUM({sheetRange3D}!D{row})";
                summaryWs.Cells[row, 5].Formula  = $"SUM({sheetRange3D}!E{row})";
                summaryWs.Cells[row, 6].Formula  = $"IFERROR(E{row}/(D{row}*1),0)";
                summaryWs.Cells[row, 7].Formula  = $"SUM({sheetRange3D}!G{row})";
                summaryWs.Cells[row, 8].Formula  = $"SUM({sheetRange3D}!H{row})";
                summaryWs.Cells[row, 9].Formula  = $"SUM({sheetRange3D}!I{row})";
                summaryWs.Cells[row, 10].Formula = $"SUM(H{row},I{row})";
                summaryWs.Cells[row, 11].Formula = $"SUM({sheetRange3D}!K{row})";
            }

            WriteTotalRowFormulas(summaryWs, totalDataRow, DataStartRow, totalDataRow - 1);

            // ひな形として使った「Template」シートは出力には不要なので削除し、
            // 「年間実績」はタブの一番最後に配置する（複製した月次シートは複製順のまま先頭側に並ぶ）。
            pkg.Workbook.Worksheets.Delete(templateWs);
            pkg.Workbook.Worksheets.MoveToEnd("年間実績");

            pkg.SaveAs(new FileInfo(outputPath));
            Logger.Info($"年間実績Excel出力完了: {outputPath}（月次シート{monthlySheetNames.Count}枚＋年間実績、車両{vehicleCount}台）");
        }

        /// <summary>
        /// [デザイン維持対応] シート内のExcelテーブルの行数を、車両数+合計行に合わせて増減させる。
        /// AddNewRows/DeleteRowsは直前の行の書式をそのまま新しい行にコピーしてしまうことがあるため、
        /// リサイズ前に「通常データ行」と「元の合計行」双方の書式を控えておき、リサイズ後に
        /// 該当する行へ正しい書式を明示的に再適用する（縞模様自体はテーブル機能側で自動描画される）。
        /// </summary>
        private static void ResizeSheetTable(ExcelWorksheet ws, int vehicleCount)
        {
            var table = ws.Tables?.FirstOrDefault();
            if (table == null) return;

            int desiredDataRows = vehicleCount + 1; // 車両行 + 合計行
            int currentDataRows = table.Address.End.Row - table.Address.Start.Row;
            int firstCol = table.Address.Start.Column;
            int lastCol  = table.Address.End.Column;
            int normalRowRef   = table.Address.Start.Row + 1; // 既存の先頭データ行（通常行の書式見本）
            int oldTotalRowRef = table.Address.End.Row;       // 既存の最終行（合計行の書式見本）

            int width = lastCol - firstCol + 1;
            var normalStyleIds = new int[width];
            var totalStyleIds  = new int[width];
            for (int col = firstCol; col <= lastCol; col++)
            {
                normalStyleIds[col - firstCol] = ws.Cells[normalRowRef, col].StyleID;
                totalStyleIds[col - firstCol]  = ws.Cells[oldTotalRowRef, col].StyleID;
            }

            if (desiredDataRows > currentDataRows)
                table.DataRows.AddNewRows(desiredDataRows - currentDataRows);
            else if (desiredDataRows < currentDataRows)
                table.DataRows.DeleteRows(desiredDataRows, currentDataRows - desiredDataRows);

            int newTotalRow = table.Address.End.Row;
            for (int row = table.Address.Start.Row + 1; row <= table.Address.End.Row; row++)
            {
                var styleSet = (row == newTotalRow) ? totalStyleIds : normalStyleIds;
                for (int col = firstCol; col <= lastCol; col++)
                    ws.Cells[row, col].StyleID = styleSet[col - firstCol];
            }
        }

        /// <summary>
        /// 車両1台分のデータ行に値・数式を書き込む（書式はResizeSheetTableで既に整えられている前提）。
        /// D列（延実在車輌数）はひな形と同じくA1/B1から月の日数を計算する数式のまま、
        /// J列（総走行ｋｍ）はテーブル名参照の代わりに単純な範囲参照に置き換えている
        /// （シート複製後もテーブル名に依存せず安全に計算されるようにするため。表示・計算結果は同じ）。
        /// </summary>
        private static void WriteVehicleRowValues(ExcelWorksheet ws, int row, int no, string shisha, string vehicleNo, MonthlyRecord record)
        {
            ws.Cells[row, 1].Value = $"№{no}";
            ws.Cells[row, 2].Value = shisha;
            ws.Cells[row, 3].Value = int.TryParse(vehicleNo, out int vn) ? (object)vn : vehicleNo;
            ws.Cells[row, 4].Formula = $"DAY(EOMONTH(DATE(VALUE(MID($A$1,2,LEN($A$1)-1))+2018,$B$1,1),0))";
            ws.Cells[row, 5].Value   = record?.JitsudouSuu ?? 0;
            ws.Cells[row, 6].Formula = $"IFERROR(E{row}/(D{row}*1),0)";
            ws.Cells[row, 7].Value   = record?.Hanso ?? 0;
            ws.Cells[row, 8].Value   = record?.YuryoKm ?? 0;
            ws.Cells[row, 9].Value   = record?.MuryoKm ?? 0;
            ws.Cells[row, 10].Formula = $"SUM(H{row}:I{row})";
            ws.Cells[row, 11].Value  = record?.Unshu ?? 0;
        }

        /// <summary>合計行の数式を書き込む（書式はResizeSheetTableで既に整えられている前提）。</summary>
        private static void WriteTotalRowFormulas(ExcelWorksheet ws, int totalRow, int firstDataRow, int lastDataRow)
        {
            ws.Cells[totalRow, 1].Value = "合\u3000\u3000\u3000計";
            ws.Cells[totalRow, 4].Formula  = $"SUM(D{firstDataRow}:D{lastDataRow})";
            ws.Cells[totalRow, 5].Formula  = $"SUM(E{firstDataRow}:E{lastDataRow})";
            ws.Cells[totalRow, 6].Formula  = $"IFERROR(E{totalRow}/(D{totalRow}*1),0)";
            ws.Cells[totalRow, 7].Formula  = $"SUM(G{firstDataRow}:G{lastDataRow})";
            ws.Cells[totalRow, 8].Formula  = $"SUM(H{firstDataRow}:H{lastDataRow})";
            ws.Cells[totalRow, 9].Formula  = $"SUM(I{firstDataRow}:I{lastDataRow})";
            ws.Cells[totalRow, 10].Formula = $"SUM(J{firstDataRow}:J{lastDataRow})";
            ws.Cells[totalRow, 11].Formula = $"SUM(K{firstDataRow}:K{lastDataRow})";
        }

        // ---- ヘルパー ----

        private static bool IsInRange(int y, int m, int sy, int sm, int ey, int em)
        {
            int val = y * 100 + m;
            return val >= sy * 100 + sm && val <= ey * 100 + em;
        }

        private static List<(int Year, int Month)> GetMonthRange(int sy, int sm, int ey, int em)
        {
            var list = new List<(int, int)>();
            int y = sy, m = sm;
            while (y * 100 + m <= ey * 100 + em)
            {
                list.Add((y, m));
                if (++m > 12) { m = 1; y++; }
            }
            return list;
        }

        private static double? GetDouble(object val)
        {
            if (val == null) return null;
            var s = val.ToString();
            if (s.StartsWith("=")) return null;
            return double.TryParse(s, out double d) ? d : null;
        }

        private static bool TryParseSheetName(string sname, out string shisha, out string vehicleNo)
        {
            shisha = null; vehicleNo = null;
            if (SheetNameMap.TryGetValue(sname, out var mapped))
            {
                shisha = mapped.Shisha; vehicleNo = mapped.VehicleNo; return true;
            }
            var m = FullSheetPattern.Match(sname);
            if (!m.Success) return false;
            shisha = m.Groups[1].Value; vehicleNo = m.Groups[2].Value; return true;
        }
    }
}
