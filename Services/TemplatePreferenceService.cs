using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace OSEMAddIn.Services
{
    internal sealed class TemplatePreferenceService
    {
        private readonly string _storePath;
        private Dictionary<string, string> _preferences;
        private readonly object _lock = new object();

        public TemplatePreferenceService()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _storePath = Path.Combine(appData, "OSEMAddIn", "template_preferences.json");
            _preferences = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Load();
        }

        public string? GetPreferredTemplate(IEnumerable<string> participants)
        {
            if (participants is null) return null;

            lock (_lock)
            {
                foreach (var participant in participants)
                {
                    if (_preferences.TryGetValue(participant, out var templateId))
                    {
                        return templateId;
                    }
                }
            }
            return null;
        }

        public void UpdatePreference(string participant, string templateId)
        {
            if (string.IsNullOrWhiteSpace(participant) || string.IsNullOrEmpty(templateId)) return;

            lock (_lock)
            {
                _preferences[participant] = templateId;
                Save();
            }
        }

        public void RemovePreference(string participant)
        {
            if (string.IsNullOrWhiteSpace(participant)) return;

            lock (_lock)
            {
                if (_preferences.Remove(participant))
                {
                    Save();
                }
            }
        }

        public IReadOnlyDictionary<string, string> GetPreferences()
        {
            lock (_lock)
            {
                return new Dictionary<string, string>(_preferences);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _preferences.Clear();
                Save();
            }
        }

        public void UpdatePreference(IEnumerable<string> participants, string templateId)
        {
            if (participants is null || string.IsNullOrEmpty(templateId)) return;

            lock (_lock)
            {
                bool changed = false;
                foreach (var participant in participants)
                {
                    if (!_preferences.TryGetValue(participant, out var current) || current != templateId)
                    {
                        _preferences[participant] = templateId;
                        changed = true;
                    }
                }

                if (changed)
                {
                    Save();
                }
            }
        }

        private void Load()
        {
            lock (_lock)
            {
                try
                {
                    if (File.Exists(_storePath))
                    {
                        var json = File.ReadAllText(_storePath);
                        _preferences = JsonConvert.DeserializeObject<Dictionary<string, string>>(json) 
                                       ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    }
                }
                catch
                {
                    // Ignore load errors
                    _preferences = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        private void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(_storePath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir!);
                }

                var json = JsonConvert.SerializeObject(_preferences, Formatting.Indented);
                File.WriteAllText(_storePath, json);
            }
            catch
            {
                // Ignore save errors
            }
        }
    }
}
