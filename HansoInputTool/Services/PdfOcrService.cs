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

        public PdfOcrService()
        {
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        }

        public async Task<List<NippoData>> AnalyzeAllPagesAsync(
            string pdfPath, string apiKey, Action<int, int> onProgress = null)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Claude APIキーが設定されていません。");
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDFが見つかりません: {pdfPath}");

            var results = new List<NippoData>();
            var pageBytes = SplitPdfToPages(pdfPath);
            Logger.Info($"PDF分割完了: {pageBytes.Count}ページ ({System.IO.Path.GetFileName(pdfPath)})");

            for (int i = 0; i < pageBytes.Count; i++)
            {
                onProgress?.Invoke(i + 1, pageBytes.Count);
                Logger.Info($"ページ {i + 1}/{pageBytes.Count} を解析中...");

                var data = await AnalyzeSinglePageAsync(pageBytes[i], apiKey);
                data.PdfPath     = pdfPath;
                data.PdfFileName = System.IO.Path.GetFileName(pdfPath);
                data.PageNumber  = i + 1;
                data.TotalPages  = pageBytes.Count;
                data.PagePdfBytes = pageBytes[i];
                results.Add(data);

                if (i < pageBytes.Count - 1)
                    await Task.Delay(500);
            }
            return results;
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
                model = "claude-haiku-4-5-20251001",
                max_tokens = 512,
                messages = new[]
                {
                    new
                    {
                        role = "user",
                        content = new object[]
                        {
                            new
                            {
                                type = "document",
                                source = new { type = "base64", media_type = "application/pdf", data = base64 }
                            },
                            new
                            {
                                type = "text",
                                text = "この自動車運転日報から以下の4項目を読み取ってください。\n\nJSONのみで返答（説明文不要）：\n{\n  \"day\": 日付の「日」のみ（数値）,\n  \"yuryo_km\": 右側「有料キロ(計)」欄の数値,\n  \"muryo_km\": 右側「無料キロ(計)」欄の数値,\n  \"shinya_minutes\": 右側「深夜作業時間」欄の分数（0または空欄なら0）,\n  \"vehicle_number\": 右上の車両番号（4桁数字）\n}\n\n注意：\n- 有料キロ・無料キロは右側集計欄の「(計)」行の数値\n- 深夜作業時間は「分記入」と書かれた欄の数値\n- 読み取れない場合はnullにする"
                            }
                        }
                    }
                }
            };

            var json = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            _httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

            var response = await _httpClient.PostAsync("https://api.anthropic.com/v1/messages", content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception($"APIエラー ({response.StatusCode}): {responseText}");

            var responseJson = JObject.Parse(responseText);
            var resultText = responseJson["content"]?[0]?["text"]?.ToString()?.Trim();

            if (string.IsNullOrEmpty(resultText))
                throw new Exception("APIからの応答が空です");

            if (resultText.Contains("{"))
            {
                var start = resultText.IndexOf('{');
                var end   = resultText.LastIndexOf('}');
                if (start >= 0 && end >= 0)
                    resultText = resultText.Substring(start, end - start + 1);
            }

            Logger.Info($"API応答(p{resultText}): {resultText}");
            return JsonConvert.DeserializeObject<NippoData>(resultText) ?? new NippoData();
        }

        public void Dispose() => _httpClient?.Dispose();
    }

    public class NippoData
    {
        [JsonProperty("day")]           public int?    Day            { get; set; }
        [JsonProperty("yuryo_km")]      public double? YuryoKm        { get; set; }
        [JsonProperty("muryo_km")]      public double? MuryoKm        { get; set; }
        [JsonProperty("shinya_minutes")]public int?    ShinyaMinutes  { get; set; }
        [JsonProperty("vehicle_number")]public string  VehicleNumber  { get; set; }

        [JsonIgnore] public string PdfPath      { get; set; }
        [JsonIgnore] public string PdfFileName  { get; set; }
        [JsonIgnore] public int    PageNumber   { get; set; }
        [JsonIgnore] public int    TotalPages   { get; set; }
        [JsonIgnore] public byte[] PagePdfBytes { get; set; }

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
