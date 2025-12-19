using System.Collections.Generic;

namespace OSEMAddIn.Models
{
    public sealed class DashboardTemplate
    {
        public string TemplateId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public List<string> Fields { get; set; } = new List<string>();
        public Dictionary<string, string> FieldRegexes { get; set; } = new Dictionary<string, string>();
        public List<string> AttachmentPaths { get; set; } = new List<string>();
    }
}
