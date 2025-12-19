using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class LlmConfigurationService
    {
        private readonly string _storePath;
        private readonly JsonSerializerSettings _serializerSettings = new()
        {
            Formatting = Formatting.Indented,
            TypeNameHandling = TypeNameHandling.None
        };

        private LlmModelConfiguration _globalConfiguration = new();
        private readonly Dictionary<string, LlmModelConfiguration> _templateOverrides = new(StringComparer.OrdinalIgnoreCase);

        public LlmConfigurationService(string? storePath = null)
        {
            _storePath = storePath ?? BuildDefaultStorePath();
            Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
            LoadFromDisk();
        }

        public LlmModelConfiguration GetGlobalConfiguration()
        {
            return Clone(_globalConfiguration);
        }

        public LlmModelConfiguration GetTemplateConfiguration(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return Clone(_globalConfiguration);
            }

            if (_templateOverrides.TryGetValue(templateId, out var value))
            {
                return Clone(value);
            }

            return Clone(_globalConfiguration);
        }

        public void SaveGlobalConfiguration(LlmModelConfiguration configuration)
        {
            if (configuration is null)
            {
                throw new ArgumentNullException(nameof(configuration));
            }

            _globalConfiguration = Clone(configuration);
            Persist();
        }

        public void SaveTemplateConfiguration(string templateId, LlmModelConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(templateId) || configuration is null)
            {
                return;
            }

            _templateOverrides[templateId] = Clone(configuration);
            Persist();
        }

        public void ClearTemplateConfiguration(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return;
            }

            if (_templateOverrides.Remove(templateId))
            {
                Persist();
            }
        }

        public LlmModelConfiguration GetEffectiveConfiguration(string templateId)
        {
            return Clone(GetTemplateConfiguration(templateId));
        }

        public bool HasTemplateConfiguration(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return false;
            }

            return _templateOverrides.ContainsKey(templateId);
        }

        private void LoadFromDisk()
        {
            if (!File.Exists(_storePath))
            {
                return;
            }

            var json = File.ReadAllText(_storePath);
            var container = JsonConvert.DeserializeObject<LlmSettingsContainer>(json, _serializerSettings);
            if (container is null)
            {
                return;
            }

            _globalConfiguration = container.Global is null ? new LlmModelConfiguration() : Clone(container.Global);
            _templateOverrides.Clear();
            foreach (var pair in container.TemplateOverrides ?? Enumerable.Empty<LlmTemplateOverride>())
            {
                if (string.IsNullOrWhiteSpace(pair.TemplateId) || pair.Configuration is null)
                {
                    continue;
                }

                _templateOverrides[pair.TemplateId] = Clone(pair.Configuration);
            }
        }

        private void Persist()
        {
            var container = new LlmSettingsContainer
            {
                Global = Clone(_globalConfiguration),
                TemplateOverrides = _templateOverrides
                    .Select(kvp => new LlmTemplateOverride
                    {
                        TemplateId = kvp.Key,
                        Configuration = Clone(kvp.Value)
                    })
                    .ToList()
            };

            var json = JsonConvert.SerializeObject(container, _serializerSettings);
            File.WriteAllText(_storePath, json);
        }

        private static LlmModelConfiguration Clone(LlmModelConfiguration configuration)
        {
            return new LlmModelConfiguration
            {
                Scope = configuration.Scope,
                Provider = configuration.Provider,
                ModelName = configuration.ModelName,
                ApiEndpoint = configuration.ApiEndpoint,
                ApiKey = configuration.ApiKey
            };
        }

        private static string BuildDefaultStorePath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(root, "OSEM", "llm-config.json");
        }

        private sealed class LlmSettingsContainer
        {
            public LlmModelConfiguration? Global { get; set; }
            public List<LlmTemplateOverride>? TemplateOverrides { get; set; }
        }

        private sealed class LlmTemplateOverride
        {
            public string TemplateId { get; set; } = string.Empty;
            public LlmModelConfiguration? Configuration { get; set; }
        }
    }
}
