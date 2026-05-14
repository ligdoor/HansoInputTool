using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace HansoInputTool.Services
{
    public class PdfOcrService : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly HttpClient _httpClient;

        // リトライ設定
        private const int MaxRetryCount  = 2;    // 失敗時の最大リトライ回数
        private const int RetryDelayMs   = 1500; // リトライ前の待機時間(ms)

        // OCR読み取りプロンプト（few-shot例付きで精度向上）
        private const string OcrPrompt = @"あなたは自動車運転日報のOCR読み取りシステムです。
添付の日報PDFから、以下の5項目を正確に読み取ってください。

【読み取り項目と場所】
1. day          : 日報の日付（「日」のみ。例: 3月15日なら 15）
2. yuryo_km     : 右側集計欄の「有料キロ」または「有料キロ(計)」の合計数値
3. muryo_km     : 右側集計欄の「無料キロ」または「無料キロ(計)」の合計数値
4. shinya_minutes: 右側集計欄の「深夜作業時間」欄の分数（記載なし・0・空欄なら 0）
5. vehicle_number: 日報右上または上部に記載された車両番号（通常4桁の数字）

【読み取りの注意点】
- 有料キロ・無料キロは、複数の走行記録がある場合は「(計)」行または最下段の合計値を使ってください
- 車両番号は「1234」のような4桁数字です。「富士吉田」「大月」などの地名は含めないでください
- 数値が読み取れない・判断できない場合は null にしてください
- 整数に見えても小数点以下がある場合は正確に読み取ってください（例: 42.5）

【出力形式】
必ずJSON形式のみで返答してください。前後の説明文・コードブロック記号（```）は不要です。
{
  ""day"": 日付の日のみ（整数）,
  ""yuryo_km"": 有料キロ合計（数値）,
  ""muryo_km"": 無料キロ合計（数値）,
  ""shinya_minutes"": 深夜作業時間（整数、なければ0）,
  ""vehicle_number"": 車両番号（文字列）
}

【出力例】
{""day"": 15, ""yuryo_km"": 42.5, ""muryo_km"": 8, ""shinya_minutes"": 0, ""vehicle_number"": ""1234""}";

        public PdfOcrService()
        {
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
        }

        public async Task<List<NippoData>> AnalyzeAllPagesAsync(
            string pdfPath, string apiKey, Action<int, int> onProgress = null)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Claude APIキーが設定されていません。");
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDFが見つかりません: {pdfPath}");

            var results   = new List<NippoData>();
            var pageBytes = SplitPdfToPages(pdfPath);
            Logger.Info($"PDF分割完了: {pageBytes.Count}ページ ({System.IO.Path.GetFileName(pdfPath)})");

            for (int i = 0; i < pageBytes.Count; i++)
            {
                onProgress?.Invoke(i + 1, pageBytes.Count);
                Logger.Info($"ページ {i + 1}/{pageBytes.Count} を解析中...");

                // リトライつきで解析
                var data = await AnalyzeWithRetryAsync(pageBytes[i], apiKey, i + 1);
                data.PdfPath      = pdfPath;
                data.PdfFileName  = System.IO.Path.GetFileName(pdfPath);
                data.PageNumber   = i + 1;
                data.TotalPages   = pageBytes.Count;
                data.PagePdfBytes = pageBytes[i];
                results.Add(data);

                if (i < pageBytes.Count - 1)
                    await Task.Delay(500);
            }
            return results;
        }

        /// <summary>
        /// 失敗した場合に最大 MaxRetryCount 回リトライする
        /// </summary>
        private async Task<NippoData> AnalyzeWithRetryAsync(byte[] pdfBytes, string apiKey, int pageNumber)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= MaxRetryCount + 1; attempt++)
            {
                try
                {
                    var data = await AnalyzeSinglePageAsync(pdfBytes, apiKey);

                    // 必須項目（日・有料キロ・無料キロ）が揃っているかチェック
                    var (isValid, missing) = data.ValidateRequired();
                    if (isValid)
                    {
                        if (attempt > 1)
                            Logger.Info($"ページ{pageNumber}: リトライ{attempt - 1}回目で成功しました。");
                        return data;
                    }

                    // 読み取り不足の場合もリトライ対象
                    lastException = new Exception($"必須項目が読み取れませんでした（不足: {missing}）");
                    Logger.Warn($"ページ{pageNumber} 試行{attempt}: {lastException.Message}");
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Logger.Warn($"ページ{pageNumber} 試行{attempt}: 例外発生 - {ex.Message}");
                }

                if (attempt <= MaxRetryCount)
                {
                    Logger.Info($"ページ{pageNumber}: {RetryDelayMs}ms後にリトライします（{attempt}/{MaxRetryCount}回目）...");
                    await Task.Delay(RetryDelayMs);
                }
            }

            // 全リトライ失敗 → 空データを返してUIで警告表示させる
            Logger.Error($"ページ{pageNumber}: {MaxRetryCount}回リトライしましたが読み取りに失敗しました。最終エラー: {lastException?.Message}");
            return new NippoData
            {
                RetryFailed  = true,
                RetryMessage = lastException?.Message ?? "不明なエラー"
            };
        }

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

        private async Task<NippoData> AnalyzeSinglePageAsync(byte[] pdfBytes, string apiKey)
        {
            var base64 = Convert.ToBase64String(pdfBytes);

            var requestBody = new
            {
                model      = "claude-haiku-4-5-20251001",
                max_tokens = 512,
                messages   = new[]
                {
                    new
                    {
                        role    = "user",
                        content = new object[]
                        {
                            new
                            {
                                type   = "document",
                                source = new { type = "base64", media_type = "application/pdf", data = base64 }
                            },
                            new
                            {
                                type = "text",
                                text = OcrPrompt
                            }
                        }
                    }
                }
            };

            var json    = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            _httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

            var response     = await _httpClient.PostAsync("https://api.anthropic.com/v1/messages", content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception($"APIエラー ({response.StatusCode}): {responseText}");

            var responseJson = JObject.Parse(responseText);
            var resultText   = responseJson["content"]?[0]?["text"]?.ToString()?.Trim();

            if (string.IsNullOrEmpty(resultText))
                throw new Exception("APIからの応答が空です");

            // JSONブロックを抽出（前後の余計なテキストを除去）
            if (resultText.Contains("{"))
            {
                var start = resultText.IndexOf('{');
                var end   = resultText.LastIndexOf('}');
                if (start >= 0 && end >= 0)
                    resultText = resultText.Substring(start, end - start + 1);
            }

            Logger.Info($"API応答: {resultText}");
            return JsonConvert.DeserializeObject<NippoData>(resultText) ?? new NippoData();
        }

        public void Dispose() => _httpClient?.Dispose();
    }

    public class NippoData
    {
        [JsonProperty("day")]            public int?    Day            { get; set; }
        [JsonProperty("yuryo_km")]       public double? YuryoKm        { get; set; }
        [JsonProperty("muryo_km")]       public double? MuryoKm        { get; set; }
        [JsonProperty("shinya_minutes")] public int?    ShinyaMinutes  { get; set; }
        [JsonProperty("vehicle_number")] public string  VehicleNumber  { get; set; }

        [JsonIgnore] public string PdfPath      { get; set; }
        [JsonIgnore] public string PdfFileName  { get; set; }
        [JsonIgnore] public int    PageNumber   { get; set; }
        [JsonIgnore] public int    TotalPages   { get; set; }
        [JsonIgnore] public byte[] PagePdfBytes { get; set; }

        /// <summary>全リトライ失敗フラグ</summary>
        [JsonIgnore] public bool   RetryFailed  { get; set; }
        /// <summary>リトライ失敗時のメッセージ</summary>
        [JsonIgnore] public string RetryMessage { get; set; }

        public (bool isValid, string missingFields) ValidateRequired()
        {
            var missing = new System.Collections.Generic.List<string>();
            if (!Day.HasValue || Day <= 0) missing.Add("日");
            if (!YuryoKm.HasValue)         missing.Add("有料キロ(計)");
            if (!MuryoKm.HasValue)         missing.Add("無料キロ(計)");
            return (missing.Count == 0, string.Join(", ", missing));
        }
    }
}
