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
    /// <summary>
    /// VehicleAnnualSummaryService の分割定義：チェック済み車両のデータ読み込み(LoadData)。
    /// </summary>
    public partial class VehicleAnnualSummaryService
    {
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
    }
}
