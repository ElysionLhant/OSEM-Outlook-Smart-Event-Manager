using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class PromptLibraryService
    {
        private readonly string _storePath;
        private List<PromptDefinition> _prompts = new();

        public event EventHandler? PromptsChanged;

        public PromptLibraryService(string? storePath = null)
        {
            _storePath = storePath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OSEMAddIn", "prompts.json");
            LoadFromDisk();
        }

        public IReadOnlyList<PromptDefinition> GetPrompts() => _prompts;

        public void AddOrUpdatePrompt(PromptDefinition prompt)
        {
            var existing = _prompts.FirstOrDefault(p => p.PromptId == prompt.PromptId);
            if (existing != null)
            {
                _prompts.Remove(existing);
            }
            _prompts.Add(prompt);
            SaveToDisk();
            PromptsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void RemovePrompt(string promptId)
        {
            var existing = _prompts.FirstOrDefault(p => p.PromptId == promptId);
            if (existing != null)
            {
                _prompts.Remove(existing);
                SaveToDisk();
                PromptsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private void LoadFromDisk()
        {
            if (File.Exists(_storePath))
            {
                try
                {
                    var json = File.ReadAllText(_storePath);
                    var loaded = JsonConvert.DeserializeObject<List<PromptDefinition>>(json);
                    if (loaded != null)
                    {
                        _prompts = loaded;
                        return;
                    }
                }
                catch
                {
                    // Ignore
                }
            }

            // Defaults
            _prompts = new List<PromptDefinition>
            {
                new PromptDefinition
                {
                    PromptId = "prompt.general",
                    DisplayName = Properties.Resources.General_Info_Extraction,
                    Body = Properties.Resources.Please_extract_dashboard_field_6914e4,
                    TemplateOverrideId = null
                }
            };
            SaveToDisk();
        }

        private void SaveToDisk()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
                var json = JsonConvert.SerializeObject(_prompts, Formatting.Indented);
                File.WriteAllText(_storePath, json);
            }
            catch
            {
                // Ignore
            }
        }
    }
}
