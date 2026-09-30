using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using HansoInputTool.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace HansoInputTool.Services
{
    public interface IPdfAiAnalyzer
    {
        Task<NippoData> AnalyzeAsync(byte[] pdfBytes, AiSettings settings);
        Task TestConnectionAsync(AiSettings settings);
    }

    public static class PdfAiAnalyzerFactory
    {
        public static IPdfAiAnalyzer Create(string provider)
        {
            return (provider ?? "Anthropic").Trim().ToLowerInvariant() switch
            {
                "openai" => new OpenAiPdfAnalyzer(),
                "gemini" => new GeminiPdfAnalyzer(),
                _ => new AnthropicPdfAnalyzer()
            };
        }
    }

    internal abstract class PdfAiAnalyzerBase : IPdfAiAnalyzer
    {
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        protected const string Prompt = @"あなたは自動車運転日報のOCR読み取りシステムです。
添付の日報PDFを帳票画像として視覚的に確認し、文字だけでなく欄の位置や丸の位置も見て以下の項目をまとめて読み取ってください。

1. day: 日報の日付の「日」のみ。例: 3月15日なら15
2. yuryo_km: 右側集計欄の「有料キロ」または「有料キロ(計)」の合計数値
3. muryo_km: 右側集計欄の「無料キロ」または「無料キロ(計)」の合計数値
4. shinya_minutes: 右側集計欄の「深夜作業時間」「深夜時間」など、深夜に明確に関連付けられた欄の分数のみ。記載なし・0・空欄なら0
5. vehicle_number: 日報上部に記載された車両番号。通常4桁の数字
6. work_type: 「搬送」「移動」のどちらに丸が付いているか。文字だけで推測せず、丸の位置で判断
7. embalming_candidate: 「エンバー」等の手書き記載があるか。候補の検出のみで登録確定ではない
8. fuel_marked: 「給油」に丸が付いているか
9. fuel_liters: 左下の「ガソリン給油数」に記載されたリッター数
10. fuel_odometer_km: 現行様式ではスタンド名の行にある「メーター指針」の数値
11. confidence: 各項目の読み取り確信度。high / medium / low のいずれか

注意:
- 有料キロ・無料キロは複数記録があっても「(計)」行または最下段の合計値を使う
- 深夜時間は「深夜作業時間」「深夜時間」など、項目名が明確に確認できる数値だけを採用する
- 「回数」「件数」「搬送回数」「数量」「人数」など、深夜時間と無関係な数値を絶対に深夜時間として使用しない
- 深夜時間の項目名や対応する数値を確認できない場合はnullとし、推測で補完しない
- 車両番号に地名は含めない
- work_typeは「搬送」「移動」の文字の近くの丸を画像上で確認する。丸なし、両方に丸、不明瞭ならnull
- 手書きエンバー記載が明確にある場合だけembalming_candidate=true、明確に無い場合はfalse、判別不能ならnull
- 給油欄に丸が明確にある場合fuel_marked=true、明確に無い場合false、判別不能ならnull
- 給油時距離はスタンド名の行に対応する「メーター指針」だけを使う。旧様式で欄がない場合や判別不能ならnull。他欄の距離から補わない
- confidenceは帳票上の文字や丸がどれだけ明瞭に見えるかについてのAI自己評価。値がnullならlow、少し曖昧ならmedium、明瞭ならhigh。値の正しさを保証するものではない
- confidenceにはday, yuryo_km, muryo_km, shinya_minutes, vehicle_number, work_type, embalming_candidate, fuel_marked, fuel_liters, fuel_odometer_kmの全キーを含める
- 判断できない値はnull。推測して埋めない
- 小数点以下がある場合は正確に読み取る
- 必ずJSONだけを返し、説明文や```は付けない

形式:
{""day"":15,""yuryo_km"":42.5,""muryo_km"":8,""shinya_minutes"":0,""vehicle_number"":""1234"",""work_type"":""搬送"",""embalming_candidate"":null,""fuel_marked"":false,""fuel_liters"":null,""fuel_odometer_km"":null,""confidence"":{""day"":""high"",""yuryo_km"":""high"",""muryo_km"":""high"",""shinya_minutes"":""medium"",""vehicle_number"":""high"",""work_type"":""high"",""embalming_candidate"":""low"",""fuel_marked"":""high"",""fuel_liters"":""low"",""fuel_odometer_km"":""low""}}";

        protected static NippoData ParseResult(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) throw new Exception("AIからの応答が空です。");
            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start >= 0 && end > start) text = text.Substring(start, end - start + 1);
            Logger.Info($"AI応答: {text}");
            return JsonConvert.DeserializeObject<NippoData>(text) ?? new NippoData();
        }

        protected static void ValidateSettings(AiSettings settings)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.ApiKey))
                throw new InvalidOperationException("AI APIキーが設定されていません。設定 → 一般設定 → AI / PDF解析設定 から設定してください。");
            if (string.IsNullOrWhiteSpace(settings.Model))
                throw new InvalidOperationException("AIモデルが設定されていません。");
        }

        public abstract Task<NippoData> AnalyzeAsync(byte[] pdfBytes, AiSettings settings);
        public abstract Task TestConnectionAsync(AiSettings settings);
    }

    internal sealed class AnthropicPdfAnalyzer : PdfAiAnalyzerBase
    {
        public override async Task<NippoData> AnalyzeAsync(byte[] pdfBytes, AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
            var request = new
            {
                model = settings.Model,
                max_tokens = 512,
                messages = new[] { new { role = "user", content = new object[]
                {
                    new { type = "document", source = new { type = "base64", media_type = "application/pdf", data = Convert.ToBase64String(pdfBytes) } },
                    new { type = "text", text = Prompt }
                } } }
            };
            client.DefaultRequestHeaders.Add("x-api-key", settings.ApiKey);
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            var response = await client.PostAsync("https://api.anthropic.com/v1/messages",
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"Anthropic APIエラー ({response.StatusCode}): {body}");
            var text = JObject.Parse(body)["content"]?[0]?["text"]?.ToString();
            return ParseResult(text);
        }

        public override async Task TestConnectionAsync(AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            var request = new { model = settings.Model, max_tokens = 16, messages = new[] { new { role = "user", content = "Reply only OK." } } };
            client.DefaultRequestHeaders.Add("x-api-key", settings.ApiKey);
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
            var response = await client.PostAsync("https://api.anthropic.com/v1/messages",
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"Anthropic APIエラー ({response.StatusCode}): {body}");
        }
    }

    internal sealed class OpenAiPdfAnalyzer : PdfAiAnalyzerBase
    {
        private const string Endpoint = "https://api.openai.com/v1/responses";

        public override async Task<NippoData> AnalyzeAsync(byte[] pdfBytes, AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", settings.ApiKey);
            var request = new
            {
                model = settings.Model,
                input = new[] { new { role = "user", content = new object[]
                {
                    new { type = "input_file", filename = "nippo.pdf", file_data = "data:application/pdf;base64," + Convert.ToBase64String(pdfBytes) },
                    new { type = "input_text", text = Prompt }
                } } }
            };
            var response = await client.PostAsync(Endpoint,
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"OpenAI APIエラー ({response.StatusCode}): {body}");
            var json = JObject.Parse(body);
            var text = json["output"]?.SelectToken("$..text")?.ToString();
            return ParseResult(text);
        }

        public override async Task TestConnectionAsync(AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", settings.ApiKey);
            var request = new { model = settings.Model, input = "Reply only OK." };
            var response = await client.PostAsync(Endpoint,
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"OpenAI APIエラー ({response.StatusCode}): {body}");
        }
    }

    internal sealed class GeminiPdfAnalyzer : PdfAiAnalyzerBase
    {
        public override async Task<NippoData> AnalyzeAsync(byte[] pdfBytes, AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
            var endpoint = $"https://generativelanguage.googleapis.com/v1beta/models/{Uri.EscapeDataString(settings.Model)}:generateContent?key={Uri.EscapeDataString(settings.ApiKey)}";
            var request = new
            {
                contents = new[] { new { parts = new object[]
                {
                    new { text = Prompt },
                    new { inlineData = new { mimeType = "application/pdf", data = Convert.ToBase64String(pdfBytes) } }
                } } }
            };
            var response = await client.PostAsync(endpoint,
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"Gemini APIエラー ({response.StatusCode}): {body}");
            var text = JObject.Parse(body)["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString();
            return ParseResult(text);
        }

        public override async Task TestConnectionAsync(AiSettings settings)
        {
            ValidateSettings(settings);
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            var endpoint = $"https://generativelanguage.googleapis.com/v1beta/models/{Uri.EscapeDataString(settings.Model)}:generateContent?key={Uri.EscapeDataString(settings.ApiKey)}";
            var request = new { contents = new[] { new { parts = new[] { new { text = "Reply only OK." } } } } };
            var response = await client.PostAsync(endpoint,
                new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json"));
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new Exception($"Gemini APIエラー ({response.StatusCode}): {body}");
        }
    }
}
