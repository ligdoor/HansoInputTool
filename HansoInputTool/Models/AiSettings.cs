namespace HansoInputTool.Models
{
    public class AiSettings
    {
        public string Provider { get; set; } = "Anthropic";
        public string Model { get; set; } = "claude-haiku-4-5-20251001";
        public string ApiKey { get; set; }
        public string Endpoint { get; set; }
    }
}
