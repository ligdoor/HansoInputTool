using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HansoInputTool.Models;
using NLog;
using OfficeOpenXml;

namespace HansoInputTool.Services
{
    /// <summary>
    /// TransferService の分割定義：通常シート（寝台車・霊柩車・CH各支社）の転記処理。
    /// ★実績月報(wsGeppo)の合計行計算など、必須ルールに関わる最重要ロジックを含む。
    /// 変更時は rules-and-learnings.md の必須ルールを必ず確認すること。
    /// </summary>
    public partial class TransferService
    {
        private void ProcessNormalSheet(ExcelPackage wbInput, ExcelPackage wbGeppo, ExcelPackage wbShukei, string sheetName, Dictionary<string, RateInfo> rates, ColumnMapping columnMap, FlagDefinitionService flagService = null, DatabaseService dbService = null, VehicleSettingsService vehicleSettingsService = null)
        {
            var wsIn = wbInput.Workbook.Worksheets[sheetName];
            var wsGeppo = wbGeppo.Workbook.Worksheets[sheetName];
            var totalRowIdx = FindTotalRow(wsIn);
            if (totalRowIdx == -1) return;

            var normalMap = columnMap.NormalSheet;
            var shukeiMap = columnMap.ShukeiSheet;

            string rateCategory = sheetName.Contains("霊柩車") ? "霊柩車" : "寝台車";
            if (!rates.TryGetValue(rateCategory, out var ratesForSheet))
            {
                Logger.Warn($"シート '{sheetName}' に対応する料金カテゴリ '{rateCategory}' が見つかりませんでした。");
                return;
            }

            // [深夜料金バグ修正] 車両設定（深夜入力方式）を優先し、未設定時のみ「大月」判定にフォールバックする。
            // これにより設定画面で深夜入力方式を「深夜時間」に変更した車両でも正しく計算されるようになる。
            bool isOotsuki = vehicleSettingsService?.IsFeeMode(sheetName) ?? sheetName.Contains("大月");
            double totalKihon = 0, totalSoko = 0, totalShinya = 0, totalSum = 0;

            // 金額ありフラグ（WithAmount）を取得して料金計算に使う
            var withAmountFlags = flagService?.Flags
                .Where(f => f.Type == FlagType.WithAmount)
                .ToList() ?? new List<FlagDefinition>();

            // ── DB使用時：DBからデータを読んでwsGeppoに書き込む ──
            if (dbService != null)
            {
                var flags   = flagService?.Flags ?? new System.Collections.ObjectModel.ReadOnlyCollection<FlagDefinition>(new List<FlagDefinition>());
                var dbRows  = dbService.GetSheetData(sheetName, flags);
                int writeRow = 3;

                // [No.3修正] DB使用時の集計値はdbRowsから直接計算する
                int dbTotalHanso = 0;
                double dbTotalYuryoKm = 0, dbTotalMuryoKm = 0;
                // [転記出力修正] 延実在車輌数／使用日数は「同じ日に何件動いても1日として1回」カウントする。
                // 同一日に複数行（複数搬送）がある場合でも重複カウントしないよう、日付の集合(HashSet)で管理する。
                var usedDaysSet = new HashSet<int>();

                foreach (var dbRow in dbRows)
                {
                    int hansoVal    = dbRow.C_Hanso ?? 0;
                    double yuryoKm  = dbRow.D_YuryoKm ?? 0;
                    double muryoKm  = dbRow.E_MuryoKm ?? 0;
                    double rowKihon = 0, rowSoko = 0, rowShinya = 0;

                    // Excelのwsに値を書き込む（geppoの行として）
                    wsGeppo.Cells[writeRow, normalMap.Day].Value        = dbRow.B_Day;
                    wsGeppo.Cells[writeRow, normalMap.HansoCount].Value = (object)hansoVal;
                    wsGeppo.Cells[writeRow, normalMap.YuryoKm].Value    = (object)yuryoKm;
                    wsGeppo.Cells[writeRow, normalMap.MuryoKm].Value    = (object)muryoKm;

                    // [No.1修正] 深夜の入力値（分 or 料金）を入力欄列に書く。
                    // ShinyaFee列への書き込みは後続の「計算後のrowShinya」のみとし、
                    // ここでは大月以外のShinyaMinutes列のみ書く。
                    if (isOotsuki)
                    {
                        // 大月は入力値をそのまま ShinyaFee 列に書く（後続のrowShinya代入と同値）
                        // ※ 後続で wsGeppo.ShinyaFee に rowShinya を書くため、ここでは書かない
                    }
                    else
                    {
                        wsGeppo.Cells[writeRow, normalMap.ShinyaMinutes].Value = (object)(dbRow.K_LateMinutes ?? 0);
                    }

                    // フラグ書き込み
                    foreach (var flag in flags)
                    {
                        int? fv = dbRow.FlagValues?.GetValueOrDefault(flag.Id);
                        wsGeppo.Cells[writeRow, flag.ExcelColumn].Value = fv == 1 ? 1 : (object)null;
                    }

                    if (hansoVal > 0)
                    {
                        rowKihon = ratesForSheet.BaseFee;
                        if (yuryoKm > 0)
                            rowSoko = (Math.Floor(yuryoKm / 10) + 1) * ratesForSheet.MileageFee;

                        foreach (var flag in withAmountFlags)
                        {
                            bool flagOn = (dbRow.FlagValues?.GetValueOrDefault(flag.Id) == 1);
                            if (!flagOn) continue;

                            bool applyBase    = flag.TargetFee == TargetFee.BaseFee || flag.TargetFee == TargetFee.Both;
                            bool applyMileage = flag.TargetFee == TargetFee.MileageFee || flag.TargetFee == TargetFee.Both;

                            if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                            {
                                if (applyBase)    rowKihon = Math.Floor(ratesForSheet.BaseFee * flag.AmountValue.Value);
                                if (applyMileage) rowSoko  = Math.Floor(rowSoko               * flag.AmountValue.Value);
                            }
                            else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                            {
                                if (applyBase)    rowKihon = flag.AmountValue.Value;
                                if (applyMileage) rowSoko  = flag.AmountValue.Value;
                            }
                        }
                        if (isOotsuki)
                            rowShinya = dbRow.H_LateFeeOotsuki ?? 0;
                        else
                        {
                            double shinyaMin = dbRow.K_LateMinutes ?? 0;
                            if (shinyaMin > 0)
                            {
                                double numBlocks   = Math.Floor(shinyaMin / 30) + 1;
                                double variableRyo = numBlocks * ratesForSheet.LateNightUnitFee;
                                rowShinya = variableRyo + ratesForSheet.LateNightFixedFee;
                            }
                        }
                    }

                    wsGeppo.Cells[writeRow, normalMap.KihonFee].Value  = (object)rowKihon;
                    wsGeppo.Cells[writeRow, normalMap.SokoFee].Value   = (object)rowSoko;
                    // [No.1修正] ShinyaFee列への書き込みはここ1か所のみ（計算済みrowShinyaを使用）
                    wsGeppo.Cells[writeRow, normalMap.ShinyaFee].Value = (object)rowShinya;
                    double rowTotal = rowKihon + rowSoko + rowShinya;
                    wsGeppo.Cells[writeRow, normalMap.TotalFee].Value  = (object)rowTotal;

                    totalKihon  += rowKihon;
                    totalSoko   += rowSoko;
                    totalShinya += rowShinya;
                    totalSum    += rowTotal;

                    // [No.3修正] dbRowsから集計値を直接計算
                    // [転記出力修正] 使用日数・延実在車輌数は搬送件数ではなく「稼働した日」の数として数える
                    if (dbRow.B_Day.HasValue) usedDaysSet.Add(dbRow.B_Day.Value);
                    dbTotalHanso  += hansoVal;
                    dbTotalYuryoKm += yuryoKm;
                    dbTotalMuryoKm += muryoKm;

                    writeRow++;
                }

                int dbTotalDays = usedDaysSet.Count;

                // [No.4修正] 実績月報（wsGeppo）103行目（合計行）にも使用・搬送・有料・無料の集計値を書き込む。
                // 合計行にはテンプレート由来のCOUNT/SUM数式が残っているが、EPPlusは保存時に再計算しないため
                // 数式のままだと未計算（空欄）になる。集計ファイルと同じ値を直接書き込んで確実に反映する。
                wsGeppo.Cells[totalRowIdx, normalMap.Day].Value        = dbTotalDays > 0 ? (object)dbTotalDays : null;
                wsGeppo.Cells[totalRowIdx, normalMap.HansoCount].Value = dbTotalHanso > 0 ? (object)dbTotalHanso : null;
                wsGeppo.Cells[totalRowIdx, normalMap.YuryoKm].Value    = dbTotalYuryoKm > 0 ? (object)dbTotalYuryoKm : null;
                wsGeppo.Cells[totalRowIdx, normalMap.MuryoKm].Value    = dbTotalMuryoKm > 0 ? (object)dbTotalMuryoKm : null;

                // [No.3修正] 集計ファイルへの書き込みをここで行い、CalculateTotals()を使わない
                // [廃車済み車両対応] 該当シートが集計ファイルに無ければ自動生成してから書き込む
                var wsShukeiForDb = EnsureShukeiSheet(wbShukei, sheetName);
                if (wsShukeiForDb != null)
                {
                    wsShukeiForDb.Cells[shukeiMap.Days].Value    = dbTotalDays;
                    wsShukeiForDb.Cells[shukeiMap.Hanso].Value   = dbTotalHanso;
                    wsShukeiForDb.Cells[shukeiMap.YuryoKm].Value = dbTotalYuryoKm;
                    wsShukeiForDb.Cells[shukeiMap.MuryoKm].Value = dbTotalMuryoKm;
                    wsShukeiForDb.Cells[shukeiMap.Total].Value   = totalSum > 0 ? totalSum : null;
                }
            }
            else
            {
            // ── Excel使用時（従来）：wsInから直接読む ──
            for (int row = 3; row < totalRowIdx; row++)
            {
                int hansoVal = GetInt(wsIn.Cells[row, normalMap.HansoCount].Value);
                double rowKihon = 0, rowSoko = 0, rowShinya = 0;

                if (hansoVal > 0)
                {
                    double yuryoKmVal = GetDouble(wsIn.Cells[row, normalMap.YuryoKm].Value);

                    // 動的フラグによる基本料金計算
                    rowKihon = ratesForSheet.BaseFee;
                    if (yuryoKmVal > 0)
                        rowSoko = (Math.Floor(yuryoKmVal / 10) + 1) * ratesForSheet.MileageFee;

                    foreach (var flag in withAmountFlags)
                    {
                        bool flagOn = GetInt(wsIn.Cells[row, flag.ExcelColumn].Value) == 1;
                        if (!flagOn) continue;

                        bool applyBase    = flag.TargetFee == TargetFee.BaseFee || flag.TargetFee == TargetFee.Both;
                        bool applyMileage = flag.TargetFee == TargetFee.MileageFee || flag.TargetFee == TargetFee.Both;

                        if (flag.AmountType == AmountType.Rate && flag.AmountValue.HasValue)
                        {
                            if (applyBase)    rowKihon = Math.Floor(ratesForSheet.BaseFee   * flag.AmountValue.Value);
                            if (applyMileage) rowSoko  = Math.Floor(rowSoko                  * flag.AmountValue.Value);
                        }
                        else if (flag.AmountType == AmountType.Fixed && flag.AmountValue.HasValue)
                        {
                            if (applyBase)    rowKihon = flag.AmountValue.Value;
                            if (applyMileage) rowSoko  = flag.AmountValue.Value;
                        }
                    }

                    if (isOotsuki)
                    {
                        rowShinya = GetDouble(wsIn.Cells[row, normalMap.ShinyaFee].Value);
                    }
                    else
                    {
                        double shinyaMin = GetDouble(wsIn.Cells[row, normalMap.ShinyaMinutes].Value);
                        if (shinyaMin > 0)
                        {
                            double numBlocks = Math.Floor(shinyaMin / 30) + 1;
                            double variableRyo = numBlocks * ratesForSheet.LateNightUnitFee;
                            rowShinya = variableRyo + ratesForSheet.LateNightFixedFee;
                        }
                    }
                }

                wsGeppo.Cells[row, normalMap.KihonFee].Value = rowKihon > 0 ? rowKihon : null;
                wsGeppo.Cells[row, normalMap.SokoFee].Value = rowSoko > 0 ? rowSoko : null;
                wsGeppo.Cells[row, normalMap.ShinyaFee].Value = rowShinya > 0 ? rowShinya : null;
                double rowTotal = rowKihon + rowSoko + rowShinya;
                wsGeppo.Cells[row, normalMap.TotalFee].Value = rowTotal > 0 ? rowTotal : null;

                totalKihon += rowKihon;
                totalSoko += rowSoko;
                totalShinya += rowShinya;
                totalSum += rowTotal;
            }
            } // end Excel使用時

            wsGeppo.Cells[totalRowIdx, normalMap.KihonFee].Value = totalKihon > 0 ? totalKihon : null;
            wsGeppo.Cells[totalRowIdx, normalMap.SokoFee].Value = totalSoko > 0 ? totalSoko : null;
            wsGeppo.Cells[totalRowIdx, normalMap.ShinyaFee].Value = totalShinya > 0 ? totalShinya : null;
            wsGeppo.Cells[totalRowIdx, normalMap.TotalFee].Value = totalSum > 0 ? totalSum : null;

            // [No.3修正] DB使用時は集計ファイルへの書き込みを上のDBブロック内で完了済み。
            // Excel使用時のみ CalculateTotals() でwsInから集計してここで書き込む。
            if (dbService == null)
            {
                var totals = CalculateTotals(wsIn, totalRowIdx, normalMap);

                // [No.4修正] 実績月報（wsGeppo）103行目（合計行）にも使用・搬送・有料・無料の集計値を書き込む。
                // 合計行にはテンプレート由来のCOUNT/SUM数式が残っているが、EPPlusは保存時に再計算しないため
                // 数式のままだと未計算（空欄）になる。集計ファイルと同じ値を直接書き込んで確実に反映する。
                wsGeppo.Cells[totalRowIdx, normalMap.Day].Value        = totals.days > 0 ? (object)totals.days : null;
                wsGeppo.Cells[totalRowIdx, normalMap.HansoCount].Value = totals.hanso > 0 ? (object)totals.hanso : null;
                wsGeppo.Cells[totalRowIdx, normalMap.YuryoKm].Value    = totals.yuryoKm > 0 ? (object)totals.yuryoKm : null;
                wsGeppo.Cells[totalRowIdx, normalMap.MuryoKm].Value    = totals.muryoKm > 0 ? (object)totals.muryoKm : null;

                // [廃車済み車両対応] 該当シートが集計ファイルに無ければ自動生成してから書き込む
                var wsShukeiForExcel = EnsureShukeiSheet(wbShukei, sheetName);
                if (wsShukeiForExcel != null)
                {
                    wsShukeiForExcel.Cells[shukeiMap.Days].Value    = totals.days;
                    wsShukeiForExcel.Cells[shukeiMap.Hanso].Value   = totals.hanso;
                    wsShukeiForExcel.Cells[shukeiMap.YuryoKm].Value = totals.yuryoKm;
                    wsShukeiForExcel.Cells[shukeiMap.MuryoKm].Value = totals.muryoKm;
                    wsShukeiForExcel.Cells[shukeiMap.Total].Value   = totalSum > 0 ? totalSum : null;
                }
            }
        }

        private (int days, int hanso, double yuryoKm, double muryoKm) CalculateTotals(ExcelWorksheet ws, int totalRowIdx, SheetColumnMap map)
        {
            int totalHanso = 0;
            double totalYuryoKm = 0, totalMuryoKm = 0;
            // [転記出力修正] 延実在車輌数／使用日数は「同じ日に何件動いても1日として1回」カウントする。
            // 同一日に複数行（複数搬送）が入力されていても重複カウントしないよう、日付の集合(HashSet)で管理する。
            var usedDaysSet = new HashSet<object>();
            for (int row = 3; row < totalRowIdx; row++)
            {
                var dayVal = ws.Cells[row, map.Day].Value;
                if (dayVal != null) usedDaysSet.Add(dayVal);
                totalHanso += GetInt(ws.Cells[row, map.HansoCount].Value);
                totalYuryoKm += GetDouble(ws.Cells[row, map.YuryoKm].Value);
                totalMuryoKm += GetDouble(ws.Cells[row, map.MuryoKm].Value);
            }
            return (usedDaysSet.Count, totalHanso, totalYuryoKm, totalMuryoKm);
        }

        private int GetInt(object val) => val == null ? 0 : (int)Convert.ToDouble(val);
        private double GetDouble(object val) => val == null ? 0.0 : Convert.ToDouble(val);
    }
}
