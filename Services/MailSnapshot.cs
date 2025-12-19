using System;
using System.Collections.Generic;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class MailSnapshot
    {
        public string EntryId { get; set; } = string.Empty;

        public string StoreId { get; set; } = string.Empty;

        public string ConversationId { get; set; } = string.Empty;

        public string InternetMessageId { get; set; } = string.Empty;

        public string Sender { get; set; } = string.Empty;

        public string To { get; set; } = string.Empty;

        public string Subject { get; set; } = string.Empty;

        public IReadOnlyCollection<string> Participants { get; set; } = Array.Empty<string>();

        public string BodyFingerprint { get; set; } = string.Empty;

        public string ThreadIndex { get; set; } = string.Empty;

    public string ThreadIndexPrefix { get; set; } = string.Empty;

        public IReadOnlyList<string> ReferenceMessageIds { get; set; } = Array.Empty<string>();

        public IReadOnlyList<string> HistoricalSubjects { get; set; } = Array.Empty<string>();

        public DateTime ReceivedOn { get; set; }

        public IReadOnlyList<AttachmentItem> Attachments { get; set; } = Array.Empty<AttachmentItem>();
    }
}
