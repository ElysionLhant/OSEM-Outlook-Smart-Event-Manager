using System;
using System.Collections.Generic;

namespace OSEMAddIn.Models
{
    internal class TemplatePackage
    {
        public PackageManifest Manifest { get; set; } = new();
        public List<DashboardTemplate> Templates { get; set; } = new();
        public List<PromptDefinition> Prompts { get; set; } = new();
        public List<PythonScriptDefinition> Scripts { get; set; } = new();
        public Dictionary<string, string> Files { get; set; } = new(); // Filename -> Base64 Content
    }

    internal class PackageManifest
    {
        public string Version { get; set; } = "1.0";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public string Description { get; set; } = string.Empty;
    }
}
