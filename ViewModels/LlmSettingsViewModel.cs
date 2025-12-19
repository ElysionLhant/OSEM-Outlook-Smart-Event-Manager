#nullable enable
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using OSEMAddIn.Models;
using OSEMAddIn.Services;

namespace OSEMAddIn.ViewModels
{
    internal sealed class LlmSettingsViewModel : ViewModelBase
    {
        private const string ScopeGlobalId = "Global";
        private const string ScopeTemplateId = "Template";
        private const string ProviderOllamaId = "Ollama";
        private const string ProviderApiId = "Api";

        private readonly LlmConfigurationService _configurationService;
        private readonly OllamaModelService _ollamaModelService;
        private readonly string? _templateId;

        private OptionItem? _selectedScope;
        private OptionItem? _selectedProvider;
        private string _modelName = string.Empty;
        private string _apiEndpoint = string.Empty;
        private string _apiKey = string.Empty;
        private bool _hasTemplateOverride;

        public LlmSettingsViewModel(LlmConfigurationService configurationService, OllamaModelService ollamaModelService, string? templateId)
        {
            _configurationService = configurationService ?? throw new ArgumentNullException(nameof(configurationService));
            _ollamaModelService = ollamaModelService ?? throw new ArgumentNullException(nameof(ollamaModelService));
            _templateId = string.IsNullOrWhiteSpace(templateId) ? null : templateId;

            ProviderOptions = new ObservableCollection<OptionItem>
            {
                new(ProviderOllamaId, Properties.Resources.Ollama_Local),
                new(ProviderApiId, "HTTP API")
            };

            ScopeOptions = new ObservableCollection<OptionItem>
            {
                new(ScopeGlobalId, Properties.Resources.Global_All_Templates)
            };

            if (_templateId is not null)
            {
                ScopeOptions.Add(new(ScopeTemplateId, Properties.Resources.Current_Template_Only));
            }

            _selectedScope = ScopeOptions.FirstOrDefault();
            LoadScopeConfiguration(_selectedScope);
            RaisePropertyChanged(nameof(SelectedScope));
            RefreshTemplateOverride();
        }

        public ObservableCollection<OptionItem> ScopeOptions { get; }
        public ObservableCollection<OptionItem> ProviderOptions { get; }
        public ObservableCollection<string> AvailableOllamaModels { get; } = new();

        public OptionItem? SelectedScope
        {
            get => _selectedScope;
            set
            {
                if (SetProperty(ref _selectedScope, value))
                {
                    LoadScopeConfiguration(value);
                    RaisePropertyChanged(nameof(IsTemplateScopeSelected));
                    RaisePropertyChanged(nameof(CanClearTemplateOverride));
                }
            }
        }

        public OptionItem? SelectedProvider
        {
            get => _selectedProvider;
            set
            {
                if (SetProperty(ref _selectedProvider, value))
                {
                    RaisePropertyChanged(nameof(IsOllama));
                    RaisePropertyChanged(nameof(ShowApiFields));
                }
            }
        }

        public string ModelName
        {
            get => _modelName;
            set => SetProperty(ref _modelName, value ?? string.Empty);
        }

        public string ApiEndpoint
        {
            get => _apiEndpoint;
            set => SetProperty(ref _apiEndpoint, value ?? string.Empty);
        }

        public string ApiKey
        {
            get => _apiKey;
            set => SetProperty(ref _apiKey, value ?? string.Empty);
        }

        public bool IsTemplateScopeAvailable => _templateId is not null;

        public bool IsTemplateScopeSelected => string.Equals(_selectedScope?.Id, ScopeTemplateId, StringComparison.Ordinal);

        public bool HasTemplateOverride
        {
            get => _hasTemplateOverride;
            private set
            {
                if (SetProperty(ref _hasTemplateOverride, value))
                {
                    RaisePropertyChanged(nameof(CanClearTemplateOverride));
                }
            }
        }

        public bool CanClearTemplateOverride => IsTemplateScopeAvailable && HasTemplateOverride;

        public bool IsOllama => string.Equals(_selectedProvider?.Id, ProviderOllamaId, StringComparison.OrdinalIgnoreCase);

        public bool ShowApiFields => !IsOllama;

        public async Task InitializeAsync()
        {
            var models = await _ollamaModelService.GetModelsAsync();
            AvailableOllamaModels.Clear();
            foreach (var model in models)
            {
                AvailableOllamaModels.Add(model);
            }

            if (!string.IsNullOrWhiteSpace(ModelName) && !AvailableOllamaModels.Any(m => string.Equals(m, ModelName, StringComparison.OrdinalIgnoreCase)))
            {
                AvailableOllamaModels.Add(ModelName);
            }
        }

        public void Save()
        {
            var provider = SelectedProvider?.Id ?? ProviderOllamaId;
            var configuration = new LlmModelConfiguration
            {
                Scope = IsTemplateScopeSelected ? "Template" : "Global",
                Provider = provider,
                ModelName = ModelName,
                ApiEndpoint = string.IsNullOrWhiteSpace(ApiEndpoint) ? null : ApiEndpoint,
                ApiKey = string.IsNullOrWhiteSpace(ApiKey) ? null : ApiKey
            };

            if (IsTemplateScopeSelected && IsTemplateScopeAvailable)
            {
                _configurationService.SaveTemplateConfiguration(_templateId!, configuration);
            }
            else
            {
                _configurationService.SaveGlobalConfiguration(configuration);
            }

            RefreshTemplateOverride();
        }

        public void ClearTemplateOverride()
        {
            if (!IsTemplateScopeAvailable)
            {
                return;
            }

            _configurationService.ClearTemplateConfiguration(_templateId!);
            RefreshTemplateOverride();
            if (IsTemplateScopeSelected)
            {
                LoadScopeConfiguration(SelectedScope);
            }
        }

        private void LoadScopeConfiguration(OptionItem? scope)
        {
            if (scope is null)
            {
                return;
            }

            LlmModelConfiguration configuration;
            if (string.Equals(scope.Id, ScopeTemplateId, StringComparison.Ordinal))
            {
                configuration = _templateId is null
                    ? _configurationService.GetGlobalConfiguration()
                    : _configurationService.GetTemplateConfiguration(_templateId);
            }
            else
            {
                configuration = _configurationService.GetGlobalConfiguration();
            }

            SelectedProvider = ProviderOptions.FirstOrDefault(p => string.Equals(p.Id, configuration.Provider, StringComparison.OrdinalIgnoreCase))
                ?? ProviderOptions.First();

            ModelName = configuration.ModelName ?? string.Empty;
            ApiEndpoint = configuration.ApiEndpoint ?? string.Empty;
            ApiKey = configuration.ApiKey ?? string.Empty;
        }

        private void RefreshTemplateOverride()
        {
            if (!IsTemplateScopeAvailable)
            {
                HasTemplateOverride = false;
                return;
            }

            HasTemplateOverride = _configurationService.HasTemplateConfiguration(_templateId!);
        }

        internal sealed class OptionItem
        {
            public OptionItem(string id, string display)
            {
                Id = id;
                Display = display;
            }

            public string Id { get; }
            public string Display { get; }

            public override string ToString() => Display;
        }
    }
}
