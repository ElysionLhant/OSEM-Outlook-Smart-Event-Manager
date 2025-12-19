namespace OSEMAddIn.Models
{
    internal sealed class AttachmentItem
    {
        public string Id { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public string SourceMailEntryId { get; set; } = string.Empty;
    }
}
