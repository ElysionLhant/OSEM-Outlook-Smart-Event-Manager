using System.Collections.Generic;

namespace OSEMAddIn.Models
{
    internal sealed class PythonScriptExecutionContext
    {
        public string EventId { get; set; } = string.Empty;
        public string EventTitle { get; set; } = string.Empty;
        public string? DashboardTemplateId { get; set; }
        public Dictionary<string, string> DashboardValues { get; set; } = new();
        public List<EmailItem> Emails { get; set; } = new();
        public List<AttachmentItem> Attachments { get; set; } = new();
    }
}
