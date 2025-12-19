namespace OSEMAddIn.Models
{
    internal sealed class PythonScriptDefinition
    {
        public string ScriptId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ScriptPath { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool IsGlobal { get; set; }
        public System.Collections.Generic.List<string> AssociatedTemplateIds { get; set; } = new();
    }
}
