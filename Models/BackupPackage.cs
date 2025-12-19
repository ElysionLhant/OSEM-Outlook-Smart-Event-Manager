using System;
using System.Collections.Generic;

namespace OSEMAddIn.Models
{
    internal class BackupPackage
    {
        public string Version { get; set; } = "1.0";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public string Description { get; set; } = string.Empty;

        public List<EventRecord> Events { get; set; } = new();
        public List<DashboardTemplate> DashboardTemplates { get; set; } = new();
        public List<EmailTemplate> EmailTemplates { get; set; } = new();
        public List<PromptDefinition> Prompts { get; set; } = new();
        public List<PythonScriptDefinition> Scripts { get; set; } = new();
        
        // Filename -> Base64 Content (for scripts and template attachments)
        public Dictionary<string, string> Files { get; set; } = new(); 
    }
}
