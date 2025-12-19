namespace OSEMAddIn.Models
{
    public sealed class LlmModelConfiguration
    {
        public string Scope { get; set; } = "Global"; // Global or Template
        public string Provider { get; set; } = "Ollama";
        public string ModelName { get; set; } = string.Empty;
        public string? ApiEndpoint { get; set; }
        public string? ApiKey { get; set; }
    }
}
