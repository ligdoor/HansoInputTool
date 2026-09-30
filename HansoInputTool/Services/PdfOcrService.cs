using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using HansoInputTool.Models;
using NLog;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace HansoInputTool.Services
{
    public class PdfOcrService : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private const int MaxRetryCount = 2;
        private const int RetryDelayMs = 1500;

        public async Task<List<NippoData>> AnalyzeAllPagesAsync(
            string pdfPath, AiSettings settings, Action<int, int> onProgress = null)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.ApiKey))
                throw new InvalidOperationException("AI APIキーが設定されていません。");
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDFが見つかりません: {pdfPath}");

            var results = new List<NippoData>();
            var pageBytes = SplitPdfToPages(pdfPath);
            var analyzer = PdfAiAnalyzerFactory.Create(settings.Provider);
            string quotaError = null;
            Logger.Info($"PDF分割完了: {pageBytes.Count}ページ ({Path.GetFileName(pdfPath)}) / Provider={settings.Provider} / Model={settings.Model}");

            for (int i = 0; i < pageBytes.Count; i++)
            {
                onProgress?.Invoke(i + 1, pageBytes.Count);
                Logger.Info($"ページ {i + 1}/{pageBytes.Count} を解析中...");

                var data = quotaError == null
                    ? await AnalyzeWithRetryAsync(analyzer, pageBytes[i], settings, i + 1)
                    : new NippoData { RetryFailed = true, RetryMessage = quotaError };
                if (quotaError == null)
                {
                    if (data.RetryFailed && IsQuotaError(data.RetryMessage))
                        quotaError = data.RetryMessage;
                }
                data.PdfPath = pdfPath;
                data.PdfFileName = Path.GetFileName(pdfPath);
                data.PageNumber = i + 1;
                data.TotalPages = pageBytes.Count;
                data.PagePdfBytes = pageBytes[i];
                results.Add(data);

                if (i < pageBytes.Count - 1) await Task.Delay(500);
            }
            return results;
        }

        private async Task<NippoData> AnalyzeWithRetryAsync(IPdfAiAnalyzer analyzer, byte[] pdfBytes, AiSettings settings, int pageNumber)
        {
            Exception lastException = null;
            for (int attempt = 1; attempt <= MaxRetryCount + 1; attempt++)
            {
                try
                {
                    var data = await analyzer.AnalyzeAsync(pdfBytes, settings);
                    var (isValid, missing) = data.ValidateCoreFields();
                    if (isValid) return data;
                    lastException = new Exception($"必須項目が読み取れませんでした（不足: {missing}）");
                    Logger.Warn($"ページ{pageNumber} 試行{attempt}: {lastException.Message}");
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Logger.Warn($"ページ{pageNumber} 試行{attempt}: {ex.Message}");
                    if (IsQuotaError(ex.Message)) break;
                }
                if (attempt <= MaxRetryCount) await Task.Delay(RetryDelayMs);
            }

            Logger.Error($"ページ{pageNumber}: リトライ失敗: {lastException?.Message}");
            return new NippoData { RetryFailed = true, RetryMessage = lastException?.Message ?? "不明なエラー" };
        }

        private static bool IsQuotaError(string message)
            => !string.IsNullOrWhiteSpace(message)
               && (message.Contains("TooManyRequests", StringComparison.OrdinalIgnoreCase)
                   || message.Contains("quota exceeded", StringComparison.OrdinalIgnoreCase)
                   || message.Contains("429", StringComparison.OrdinalIgnoreCase));

        private static List<byte[]> SplitPdfToPages(string pdfPath)
        {
            var pages = new List<byte[]>();
            using var srcDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
            for (int i = 0; i < srcDoc.PageCount; i++)
            {
                using var singleDoc = new PdfDocument();
                singleDoc.AddPage(srcDoc.Pages[i]);
                using var ms = new MemoryStream();
                singleDoc.Save(ms);
                pages.Add(ms.ToArray());
            }
            return pages;
        }

        public void Dispose() { }
    }

    public class NippoData
    {
        [Newtonsoft.Json.JsonProperty("day")] public int? Day { get; set; }
        [Newtonsoft.Json.JsonProperty("yuryo_km")] public double? YuryoKm { get; set; }
        [Newtonsoft.Json.JsonProperty("muryo_km")] public double? MuryoKm { get; set; }
        [Newtonsoft.Json.JsonProperty("shinya_minutes")] public int? ShinyaMinutes { get; set; }
        [Newtonsoft.Json.JsonProperty("vehicle_number")] public string VehicleNumber { get; set; }
        [Newtonsoft.Json.JsonProperty("work_type")] public string WorkType { get; set; }
        [Newtonsoft.Json.JsonProperty("embalming_candidate")] public bool? EmbalmingCandidate { get; set; }
        [Newtonsoft.Json.JsonProperty("fuel_marked")] public bool? FuelMarked { get; set; }
        [Newtonsoft.Json.JsonProperty("fuel_liters")] public double? FuelLiters { get; set; }
        [Newtonsoft.Json.JsonProperty("fuel_odometer_km")] public double? FuelOdometerKm { get; set; }
        [Newtonsoft.Json.JsonProperty("confidence")] public Dictionary<string, string> Confidence { get; set; }
        [Newtonsoft.Json.JsonIgnore] public string PdfPath { get; set; }
        [Newtonsoft.Json.JsonIgnore] public string PdfFileName { get; set; }
        [Newtonsoft.Json.JsonIgnore] public int PageNumber { get; set; }
        [Newtonsoft.Json.JsonIgnore] public int TotalPages { get; set; }
        [Newtonsoft.Json.JsonIgnore] public byte[] PagePdfBytes { get; set; }
        [Newtonsoft.Json.JsonIgnore] public bool RetryFailed { get; set; }
        [Newtonsoft.Json.JsonIgnore] public string RetryMessage { get; set; }

        public (bool isValid, string missingFields) ValidateRequired()
        {
            var missing = new List<string>();
            if (!Day.HasValue || Day <= 0) missing.Add("日");
            if (!YuryoKm.HasValue) missing.Add("有料キロ(計)");
            if (!MuryoKm.HasValue) missing.Add("無料キロ(計)");
            if (WorkType != "搬送" && WorkType != "移動") missing.Add("搬送/移動の丸");
            if (!FuelMarked.HasValue) missing.Add("給油の丸");
            if (FuelMarked == true && (!FuelLiters.HasValue || !FuelOdometerKm.HasValue)) missing.Add("給油リッター/給油時距離");
            return (missing.Count == 0, string.Join(", ", missing));
        }

        public (bool isValid, string missingFields) ValidateCoreFields()
        {
            var missing = new List<string>();
            if (!Day.HasValue || Day <= 0) missing.Add("日");
            if (!YuryoKm.HasValue) missing.Add("有料キロ(計)");
            if (!MuryoKm.HasValue) missing.Add("無料キロ(計)");
            return (missing.Count == 0, string.Join(", ", missing));
        }
    }
}
