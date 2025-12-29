#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;

namespace OSEMAddIn.Models
{
    internal sealed class EventRecord
    {
        public string EventId { get; set; } = string.Empty;
        public string EventTitle { get; set; } = string.Empty;
        public string DisplayColumnSource { get; set; } = string.Empty;
        public string DisplayColumnCustomValue { get; set; } = string.Empty;
        public EventStatus Status { get; set; } = EventStatus.Open;
        public int PriorityLevel { get; set; } = 0;
        public string DashboardTemplateId { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        public DateTime LastUpdatedOn { get; set; } = DateTime.UtcNow;
        public List<string> ConversationIds { get; set; } = new();
        public List<DashboardItem> DashboardItems { get; set; } = new();
        public List<EmailItem> Emails { get; set; } = new();
        public List<AttachmentItem> Attachments { get; set; } = new();
        public List<string> AdditionalFiles { get; set; } = new();
        public List<string> ExcludedTemplateFiles { get; set; } = new();
        public HashSet<string> RelatedSubjects { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Participants { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> ProcessedMessageIds { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> NotFoundMessageIds { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public string LocalFolderPath { get; set; } = string.Empty;

        public bool IsConversationTracked(string conversationId) => ConversationIds.Contains(conversationId, StringComparer.OrdinalIgnoreCase);
    }
}
