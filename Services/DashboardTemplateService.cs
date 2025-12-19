using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class DashboardTemplateService
    {
        private readonly string _storePath;
        private List<DashboardTemplate> _templates = new();

        public event EventHandler? TemplatesChanged;

        public DashboardTemplateService(string? storePath = null)
        {
            _storePath = storePath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OSEMAddIn", "dashboard_templates.json");
            LoadFromDisk();
        }

        public IReadOnlyList<DashboardTemplate> GetTemplates() => _templates;

        public DashboardTemplate? FindById(string? templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return null;
            }

            return _templates.FirstOrDefault(t => t.TemplateId == templateId);
        }

        public void AddOrUpdateTemplate(DashboardTemplate template)
        {
            var existing = _templates.FirstOrDefault(t => t.TemplateId == template.TemplateId);
            if (existing != null)
            {
                _templates.Remove(existing);
            }
            _templates.Add(template);
            SaveToDisk();
            TemplatesChanged?.Invoke(this, EventArgs.Empty);
        }

        public void RemoveTemplate(string templateId)
        {
            var existing = _templates.FirstOrDefault(t => t.TemplateId == templateId);
            if (existing != null)
            {
                _templates.Remove(existing);
                SaveToDisk();
                TemplatesChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private void LoadFromDisk()
        {
            if (File.Exists(_storePath))
            {
                try
                {
                    var json = File.ReadAllText(_storePath);
                    var loaded = JsonConvert.DeserializeObject<List<DashboardTemplate>>(json);
                    if (loaded != null)
                    {
                        foreach (var t in loaded)
                        {
                            if (t.Fields == null) t.Fields = new List<string>();
                            if (t.FieldRegexes == null) t.FieldRegexes = new Dictionary<string, string>();
                            if (t.AttachmentPaths == null) t.AttachmentPaths = new List<string>();
                        }
                        
                        // Deduplicate by TemplateId, keeping the last one or first one. 
                        // Here we keep the first occurrence to be safe.
                        _templates = loaded
                            .GroupBy(t => t.TemplateId)
                            .Select(g => g.First())
                            .ToList();
                        return;
                    }
                }
                catch
                {
                    // Ignore errors, fallback to defaults
                }
            }

            // Defaults
            _templates = new List<DashboardTemplate>
            {
                new DashboardTemplate
                {
                    TemplateId = "GEN",
                    DisplayName = Properties.Resources.General_Template,
                    Description = Properties.Resources.Basic_fields_for_general_events,
                    Fields = new List<string> { "Title", "Owner", "DueDate", "Notes" }
                },
                new DashboardTemplate
                {
                    TemplateId = "LOGISTICS_DEMO",
                    DisplayName = Properties.Resources.Logistics_Demo_Template,
                    Description = Properties.Resources.Example_dashboard_fields_for_t_7331d7,
                    Fields = new List<string> { "Shipper", "Terms", "Routing", "ETD", "Flight", "MAWB", "HAWB" },
                    FieldRegexes = new Dictionary<string, string>
                    {
                        ["HAWB"] = @"HAWB[:\s-]*(?<value>[A-Z0-9]{8,})",
                        ["Flight"] = @"Flight[:\s-]*(?<value>[A-Z0-9]{3,}\s?[0-9]{1,4})"
                    }
                }
            };
            SaveToDisk();
        }

        private void SaveToDisk()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
                var json = JsonConvert.SerializeObject(_templates, Formatting.Indented);
                File.WriteAllText(_storePath, json);
            }
            catch
            {
                // Ignore save errors
            }
        }
    }
}
