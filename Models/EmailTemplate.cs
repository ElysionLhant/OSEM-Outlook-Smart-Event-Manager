using System;

namespace OSEMAddIn.Models
{
    internal sealed class EmailTemplate
    {
        public string TemplateId { get; set; } = Guid.NewGuid().ToString("N");
        public string DisplayName { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public EmailTemplateType TemplateType { get; set; } = EmailTemplateType.Compose;
    }
}
