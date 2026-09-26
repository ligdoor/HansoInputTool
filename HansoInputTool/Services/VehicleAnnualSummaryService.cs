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
    public partial class VehicleAnnualSummaryService
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

    }
}
