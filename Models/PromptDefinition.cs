namespace OSEMAddIn.Models
{
    public sealed class PromptDefinition
    {
        public string PromptId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string? TemplateOverrideId { get; set; }
    }
}
