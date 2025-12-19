using System;

namespace OSEMAddIn.Models
{
    internal sealed class EmailItem
    {
        public string EntryId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string InternetMessageId { get; set; } = string.Empty;
        public string Sender { get; set; } = string.Empty;
        public string To { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string[] Participants { get; set; } = Array.Empty<string>();
        public string BodyFingerprint { get; set; } = string.Empty;
        public string ThreadIndex { get; set; } = string.Empty;
        public string ThreadIndexPrefix { get; set; } = string.Empty;
        public string[] ReferenceMessageIds { get; set; } = Array.Empty<string>();
        public DateTime ReceivedOn { get; set; }
        public bool IsNewOrUpdated { get; set; }
        public bool IsRemoved { get; set; }
    }
}
