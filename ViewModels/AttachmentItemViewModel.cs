using System;
using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    internal sealed class AttachmentItemViewModel
    {
        public AttachmentItemViewModel(AttachmentItem attachment, EmailItem? sourceEmail)
        {
            Id = attachment.Id;
            FileName = attachment.FileName;
            FileType = attachment.FileType;
            FileSizeBytes = attachment.FileSizeBytes;
            SourceMailEntryId = attachment.SourceMailEntryId;

            if (sourceEmail != null)
            {
                Sender = sourceEmail.Sender;
                Subject = sourceEmail.Subject;
                ReceivedOn = sourceEmail.ReceivedOn.ToLocalTime();
            }
        }

        public string Id { get; }
        public string FileName { get; }
        public string FileType { get; }
        public long FileSizeBytes { get; }
        public string SourceMailEntryId { get; }
        public string Sender { get; } = string.Empty;
        public string Subject { get; } = string.Empty;
        public DateTime? ReceivedOn { get; }

        public int SortPriority
        {
            get
            {
                if (string.IsNullOrWhiteSpace(FileType)) return 10;
                var ext = FileType.TrimStart('.').ToLowerInvariant();
                return ext switch
                {
                    "doc" or "docx" or "xls" or "xlsx" or "pdf" or "ppt" or "pptx" => 0,
                    "txt" or "csv" or "xml" or "json" => 5,
                    _ => 10
                };
            }
        }
    }
}
