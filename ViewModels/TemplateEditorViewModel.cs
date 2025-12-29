#nullable enable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;
using OSEMAddIn.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn.ViewModels
{
    internal sealed class TemplateEditorViewModel : ViewModelBase
    {
        private readonly ServiceContainer _services;
        private DashboardTemplate? _selectedTemplate;
        private PromptDefinition? _selectedPrompt;
        private string _newTemplateName = string.Empty;
        private LlmModelConfiguration _llmConfiguration;

        public TemplateEditorViewModel(ServiceContainer services)
        {
            _services = services ?? throw new ArgumentNullException(nameof(services));
            
            Templates = new ObservableCollection<DashboardTemplate>(_services.DashboardTemplates.GetTemplates());
            Prompts = new ObservableCollection<PromptDefinition>(_services.PromptLibrary.GetPrompts());
            Scripts = new ObservableCollection<PythonScriptDefinition>(_services.PythonScripts.DiscoverScripts());
            _llmConfiguration = _services.LlmConfigurations.GetGlobalConfiguration();

            AddTemplateCommand = new RelayCommand(_ => AddTemplate());
            CopyTemplateCommand = new RelayCommand(_ => CopyTemplate(), _ => SelectedTemplate != null);
            RemoveTemplateCommand = new RelayCommand(_ => RemoveTemplate(), _ => SelectedTemplate != null);
            ImportTemplateCommand = new RelayCommand(_ => ImportTemplate());
            ExportSelectedTemplatesCommand = new RelayCommand(param => ExportSelectedTemplates(param));
            ExportAllTemplatesCommand = new RelayCommand(_ => ExportAllTemplates());

            AddFieldCommand = new RelayCommand(_ => AddField(), _ => SelectedTemplate != null);
            RemoveFieldCommand = new RelayCommand(param => RemoveField(param as string), _ => SelectedTemplate != null);
            
            AddPromptCommand = new RelayCommand(_ => AddPrompt());
            RemovePromptCommand = new RelayCommand(_ => RemovePrompt(), _ => SelectedPrompt != null);
            SavePromptCommand = new RelayCommand(_ => SavePrompt(), _ => SelectedPrompt != null);

            AddScriptCommand = new RelayCommand(_ => AddScript());
            RemoveScriptCommand = new RelayCommand(_ => RemoveScript(), _ => SelectedScript != null);
            SaveScriptCommand = new RelayCommand(_ => SaveScript(), _ => SelectedScript != null);

            SaveLlmConfigCommand = new RelayCommand(_ => SaveLlmConfig(true));
            
            ExportBackupCommand = new AsyncRelayCommand(_ => ExportBackupAsync());
            ImportBackupCommand = new AsyncRelayCommand(_ => ImportBackupAsync());

            AddFileCommand = new RelayCommand(_ => AddFile(), _ => SelectedTemplate != null);
            RemoveFileCommand = new RelayCommand(param => RemoveFile(param as string), _ => SelectedTemplate != null);
            
            SaveAllCommand = new RelayCommand(_ => SaveAll());

            BrowseStoragePathCommand = new RelayCommand(_ => BrowseStoragePath());

            AvailableVariables = new ObservableCollection<PromptVariable>
            {
                new PromptVariable { Name = Properties.Resources.Dashboard_Structure, Description = Properties.Resources.Auto_generated_JSON_field_stru_56821c, InsertText = "{{DASHBOARD_JSON}}" },
                new PromptVariable { Name = Properties.Resources.Email_Subject, Description = Properties.Resources.Subject_of_the_currently_selected_email, InsertText = "{{MailSubject}}" },
                new PromptVariable { Name = Properties.Resources.Sender, Description = Properties.Resources.Sender_name_of_the_currently_selected_email, InsertText = "{{MailSender}}" },
                new PromptVariable { Name = Properties.Resources.Email_Body_Full, Description = Properties.Resources.Full_body_content_of_the_curre_6337dd, InsertText = "{{MAIL_BODY}}" },
                new PromptVariable { Name = Properties.Resources.Email_Body_Latest_Only, Description = Properties.Resources.Attempts_to_extract_only_the_l_418b31, InsertText = "{{MAIL_BODY_LATEST}}" }
            };

            SelectableTemplatesView = CollectionViewSource.GetDefaultView(_allSelectableTemplates);
            SelectableTemplatesView.Filter = FilterTemplates;

            SelectableTemplatesForScriptView = CollectionViewSource.GetDefaultView(_allSelectableTemplatesForScript);
            SelectableTemplatesForScriptView.Filter = FilterScriptTemplates;

            AddFolderCommand = new RelayCommand(_ => AddFolder());
            RemoveFolderCommand = new RelayCommand(param => RemoveFolder(param as MonitoredFolderViewModel));
            LoadMonitoredFolders();

            if (IsOllamaProvider)
            {
                RefreshOllamaModels();
            }
        }

        public ObservableCollection<DashboardTemplate> Templates { get; }
        public ObservableCollection<PromptDefinition> Prompts { get; }
        public ObservableCollection<PythonScriptDefinition> Scripts { get; }
        public ObservableCollection<PromptVariable> AvailableVariables { get; }
        public ObservableCollection<string> OllamaModels { get; } = new ObservableCollection<string>();

        public ICommand ImportTemplateCommand { get; }
        public ICommand ExportSelectedTemplatesCommand { get; }
        public ICommand ExportAllTemplatesCommand { get; }
        public ICommand ExportBackupCommand { get; }
        public ICommand ImportBackupCommand { get; }
        public ICommand BrowseStoragePathCommand { get; }

        private ObservableCollection<SelectableTemplateViewModel> _allSelectableTemplates = new();
        public ICollectionView SelectableTemplatesView { get; }

        private ObservableCollection<SelectableTemplateViewModel> _allSelectableTemplatesForScript = new();
        public ICollectionView SelectableTemplatesForScriptView { get; }

        private PythonScriptDefinition? _selectedScript;
        private string _templateSearchText = string.Empty;
        public string TemplateSearchText
        {
            get => _templateSearchText;
            set
            {
                if (SetProperty(ref _templateSearchText, value))
                {
                    SelectableTemplatesView.Refresh();
                }
            }
        }

        public string SelectedTemplatesSummary
        {
            get
            {
                var selected = _allSelectableTemplates.Where(t => t.IsSelected).Select(t => t.DisplayName).ToList();
                if (selected.Count == 0) return Properties.Resources.No_template_associated_will_apply_to_all_templates;
                return Properties.Resources.Associated + string.Join(", ", selected);
            }
        }

        public DashboardTemplate? SelectedTemplate
        {
            get => _selectedTemplate;
            set
            {
                if (_selectedTemplate == value) return;
                _selectedTemplate = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(IsDetailVisible));
                RaisePropertyChanged(nameof(TemplateFields));
                UpdateTemplateRegexes();
                RaisePropertyChanged(nameof(TemplateFiles));
                RemoveTemplateCommand.RaiseCanExecuteChanged();
                CopyTemplateCommand.RaiseCanExecuteChanged();
                AddFieldCommand.RaiseCanExecuteChanged();
                AddFileCommand.RaiseCanExecuteChanged();
                RemoveFileCommand.RaiseCanExecuteChanged();
            }
        }

        public PromptDefinition? SelectedPrompt
        {
            get => _selectedPrompt;
            set
            {
                if (_selectedPrompt == value) return;
                _selectedPrompt = value;
                RaisePropertyChanged();
                UpdateSelectableTemplates();
                ((RelayCommand)RemovePromptCommand).RaiseCanExecuteChanged();
                ((RelayCommand)SavePromptCommand).RaiseCanExecuteChanged();
            }
        }

        public PythonScriptDefinition? SelectedScript
        {
            get => _selectedScript;
            set
            {
                if (_selectedScript == value) return;
                _selectedScript = value;
                RaisePropertyChanged();
                UpdateScriptTemplateSelection();
                ((RelayCommand)RemoveScriptCommand).RaiseCanExecuteChanged();
                ((RelayCommand)SaveScriptCommand).RaiseCanExecuteChanged();
            }
        }

        private string _scriptTemplateSearchText = string.Empty;
        public string ScriptTemplateSearchText
        {
            get => _scriptTemplateSearchText;
            set
            {
                if (SetProperty(ref _scriptTemplateSearchText, value))
                {
                    SelectableTemplatesForScriptView.Refresh();
                }
            }
        }

        public ICommand AddScriptCommand { get; }
        public ICommand RemoveScriptCommand { get; }
        public ICommand SaveScriptCommand { get; }

        public string NewTemplateName
        {
            get => _newTemplateName;
            set
            {
                if (_newTemplateName == value) return;
                _newTemplateName = value;
                RaisePropertyChanged();
                AddTemplateCommand.RaiseCanExecuteChanged();
            }
        }

        public LlmModelConfiguration LlmConfiguration
        {
            get => _llmConfiguration;
            set
            {
                if (_llmConfiguration == value) return;
                _llmConfiguration = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(LlmProvider));
                RaisePropertyChanged(nameof(IsOllamaProvider));
                RaisePropertyChanged(nameof(IsCustomProvider));
            }
        }

        public string LlmProvider
        {
            get => LlmConfiguration?.Provider ?? "Ollama";
            set
            {
                if (LlmConfiguration != null && LlmConfiguration.Provider != value)
                {
                    LlmConfiguration.Provider = value;
                    RaisePropertyChanged();
                    RaisePropertyChanged(nameof(IsOllamaProvider));
                    RaisePropertyChanged(nameof(IsCustomProvider));
                    
                    if (IsOllamaProvider)
                    {
                        RefreshOllamaModels();
                    }
                }
            }
        }

        public bool IsOllamaProvider => LlmProvider == "Ollama";
        public bool IsCustomProvider => LlmProvider == "Custom";

        private bool _isMultiSelection;
        public bool IsMultiSelection
        {
            get => _isMultiSelection;
            set
            {
                if (_isMultiSelection == value) return;
                _isMultiSelection = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(IsDetailVisible));
            }
        }

        public bool IsDetailVisible => SelectedTemplate != null && !IsMultiSelection;

        // Helper properties for binding
        public ObservableCollection<string> TemplateFields => SelectedTemplate != null ? new ObservableCollection<string>(SelectedTemplate.Fields) : new ObservableCollection<string>();
        
        private ObservableCollection<RegexEntry> _templateRegexes = new ObservableCollection<RegexEntry>();
        private RegexEntry? _selectedRegexEntry;

        public ObservableCollection<RegexEntry> TemplateRegexes 
        {
            get => _templateRegexes;
            set
            {
                if (_templateRegexes == value) return;
                _templateRegexes = value;
                RaisePropertyChanged();
            }
        }

        public RegexEntry? SelectedRegexEntry
        {
            get => _selectedRegexEntry;
            set
            {
                if (_selectedRegexEntry == value) return;
                _selectedRegexEntry = value;
                RaisePropertyChanged();
            }
        }

        private void UpdateTemplateRegexes()
        {
            if (SelectedTemplate == null) 
            {
                TemplateRegexes = new ObservableCollection<RegexEntry>();
                return;
            }

            var list = new ObservableCollection<RegexEntry>();
            if (SelectedTemplate.FieldRegexes == null) SelectedTemplate.FieldRegexes = new Dictionary<string, string>();
            
            foreach (var field in SelectedTemplate.Fields)
            {
                var regex = SelectedTemplate.FieldRegexes.ContainsKey(field) ? SelectedTemplate.FieldRegexes[field] : string.Empty;
                var entry = new RegexEntry { FieldName = field, RegexPattern = regex };
                SetupRegexEntry(entry);
                list.Add(entry);
            }
            TemplateRegexes = list;
        }

        private void SetupRegexEntry(RegexEntry entry)
        {
             entry.PropertyChanged += (s, e) => 
             {
                 if (SelectedTemplate == null) return;
                 if (e.PropertyName == nameof(RegexEntry.RegexPattern))
                 {
                     if (SelectedTemplate.FieldRegexes.ContainsKey(entry.FieldName))
                         SelectedTemplate.FieldRegexes[entry.FieldName] = entry.RegexPattern;
                     else
                         SelectedTemplate.FieldRegexes.Add(entry.FieldName, entry.RegexPattern);
                 }
             };

             entry.OnRename = (oldName, newName) => 
             {
                 if (SelectedTemplate == null) return;
                 var index = SelectedTemplate.Fields.IndexOf(oldName);
                 if (index != -1)
                 {
                     SelectedTemplate.Fields[index] = newName;
                 }

                 if (SelectedTemplate.FieldRegexes.ContainsKey(oldName))
                 {
                     var val = SelectedTemplate.FieldRegexes[oldName];
                     SelectedTemplate.FieldRegexes.Remove(oldName);
                     SelectedTemplate.FieldRegexes[newName] = val;
                 }
                 
                 _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                 RaisePropertyChanged(nameof(TemplateFields));
             };
        }

        private string? _selectedFile;
        public string? SelectedFile
        {
            get => _selectedFile;
            set
            {
                if (_selectedFile == value) return;
                _selectedFile = value;
                RaisePropertyChanged();
            }
        }

        public ObservableCollection<string> TemplateFiles => SelectedTemplate != null ? new ObservableCollection<string>(SelectedTemplate.AttachmentPaths ?? new List<string>()) : new ObservableCollection<string>();

        public RelayCommand AddTemplateCommand { get; }
        public RelayCommand CopyTemplateCommand { get; }
        public RelayCommand RemoveTemplateCommand { get; }
        public RelayCommand AddFieldCommand { get; }
        public RelayCommand RemoveFieldCommand { get; }
        public ICommand AddPromptCommand { get; }
        public ICommand RemovePromptCommand { get; }
        public ICommand SavePromptCommand { get; }
        public ICommand SaveLlmConfigCommand { get; }
        public RelayCommand AddFileCommand { get; }
        public RelayCommand RemoveFileCommand { get; }
        public ICommand SaveAllCommand { get; }

        public ObservableCollection<MonitoredFolderViewModel> MonitoredFolders { get; } = new ObservableCollection<MonitoredFolderViewModel>();
        public ICommand AddFolderCommand { get; }
        public ICommand RemoveFolderCommand { get; }

        private void AddTemplate()
        {
            var name = Properties.Resources.New_Template;
            var newTemplate = new DashboardTemplate
            {
                TemplateId = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant(),
                DisplayName = name,
                Fields = new List<string> { "Title", "Notes" }
            };
            _services.DashboardTemplates.AddOrUpdateTemplate(newTemplate);
            Templates.Add(newTemplate);
            SelectedTemplate = newTemplate;
        }

        private void CopyTemplate()
        {
            if (SelectedTemplate == null) return;

            var oldTemplateId = SelectedTemplate.TemplateId;

            var baseName = SelectedTemplate.DisplayName;
            var newName = baseName;
            int i = 1;
            while (Templates.Any(t => t.DisplayName == newName))
            {
                newName = $"{baseName}{i++}";
            }

            var newTemplate = new DashboardTemplate
            {
                TemplateId = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant(),
                DisplayName = newName,
                Description = SelectedTemplate.Description,
                Fields = new List<string>(SelectedTemplate.Fields),
                FieldRegexes = new Dictionary<string, string>(SelectedTemplate.FieldRegexes ?? new Dictionary<string, string>()),
                AttachmentPaths = new List<string>(SelectedTemplate.AttachmentPaths ?? new List<string>())
            };

            _services.DashboardTemplates.AddOrUpdateTemplate(newTemplate);
            Templates.Add(newTemplate);

            // 1. Associate Prompts
            foreach (var prompt in _services.PromptLibrary.GetPrompts().ToList())
            {
                if (!string.IsNullOrEmpty(prompt.TemplateOverrideId))
                {
                    var ids = prompt.TemplateOverrideId!.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(id => id.Trim())
                        .ToList();
                    
                    if (ids.Contains(oldTemplateId))
                    {
                        ids.Add(newTemplate.TemplateId);
                        prompt.TemplateOverrideId = string.Join(",", ids);
                        _services.PromptLibrary.AddOrUpdatePrompt(prompt);
                    }
                }
            }

            // 2. Associate Scripts
            foreach (var script in _services.PythonScripts.DiscoverScripts().ToList())
            {
                if (script.AssociatedTemplateIds != null && script.AssociatedTemplateIds.Contains(oldTemplateId))
                {
                    if (!script.AssociatedTemplateIds.Contains(newTemplate.TemplateId))
                    {
                        script.AssociatedTemplateIds.Add(newTemplate.TemplateId);
                        _services.PythonScripts.UpdateScriptMetadata(script);
                    }
                }
            }

            SelectedTemplate = newTemplate;
        }

        private void RemoveTemplate()
        {
            if (SelectedTemplate == null) return;
            _services.DashboardTemplates.RemoveTemplate(SelectedTemplate.TemplateId);
            Templates.Remove(SelectedTemplate);
            SelectedTemplate = null;
        }

        private void ImportTemplate()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "OSEM Package (*.osempack)|*.osempack",
                Title = Properties.Resources.Import_Template_Package
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _services.TemplatePackages.ImportPackage(dialog.FileName);
                    
                    // Refresh
                    Templates.Clear();
                    foreach (var t in _services.DashboardTemplates.GetTemplates()) Templates.Add(t);
                    
                    Prompts.Clear();
                    foreach (var p in _services.PromptLibrary.GetPrompts()) Prompts.Add(p);
                    
                    Scripts.Clear();
                    foreach (var s in _services.PythonScripts.DiscoverScripts()) Scripts.Add(s);
                    
                    MessageBox.Show(Properties.Resources.Import_successful, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(Properties.Resources.Import_failed_ex_Message, ex.Message), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportSelectedTemplates(object? param)
        {
            var selectedItems = (param as System.Collections.IList)?.Cast<DashboardTemplate>().ToList();
            if (selectedItems == null || selectedItems.Count == 0)
            {
                MessageBox.Show(Properties.Resources.Please_select_a_template_to_export_first, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "OSEM Package (*.osempack)|*.osempack",
                Title = Properties.Resources.Export_Template_Package,
                FileName = selectedItems.Count == 1 ? $"{selectedItems[0].DisplayName}.osempack" : "templates.osempack"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _services.TemplatePackages.ExportPackage(dialog.FileName, selectedItems);
                    MessageBox.Show(Properties.Resources.Export_successful, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportAllTemplates()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "OSEM Package (*.osempack)|*.osempack",
                Title = Properties.Resources.Export_All_Templates,
                FileName = "all_templates.osempack"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _services.TemplatePackages.ExportPackage(dialog.FileName, Templates.ToList());
                    MessageBox.Show(Properties.Resources.Export_successful, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void AddField()
        {
            if (SelectedTemplate == null) return;
            
            string baseName = "New Field";
            string newName = baseName;
            int i = 1;
            while (SelectedTemplate.Fields.Contains(newName))
            {
                newName = $"{baseName} {i++}";
            }

            // We need to modify the underlying list and notify
            var list = new List<string>(SelectedTemplate.Fields);
            list.Add(newName);
            SelectedTemplate.Fields = list;
            
            _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
            
            var entry = new RegexEntry { FieldName = newName, RegexPattern = "" };
            SetupRegexEntry(entry);
            TemplateRegexes.Add(entry);
            SelectedRegexEntry = entry;
            
            RaisePropertyChanged(nameof(TemplateFields));
        }

        private void RemoveField(string? fieldName)
        {
            if (SelectedTemplate == null || fieldName == null) return;
            
            var list = new List<string>(SelectedTemplate.Fields);
            if (list.Remove(fieldName))
            {
                SelectedTemplate.Fields = list;
                if (SelectedTemplate.FieldRegexes.ContainsKey(fieldName))
                {
                    SelectedTemplate.FieldRegexes.Remove(fieldName);
                }
                _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                
                var entry = TemplateRegexes.FirstOrDefault(x => x.FieldName == fieldName);
                if (entry != null)
                {
                    TemplateRegexes.Remove(entry);
                }

                RaisePropertyChanged(nameof(TemplateFields));
            }
        }

        private void AddPrompt()
        {
            var newPrompt = new PromptDefinition
            {
                PromptId = "prompt." + Guid.NewGuid().ToString("N").Substring(0, 6),
                DisplayName = "New Prompt",
                Body = "请根据以下邮件内容提取仪表盘字段：{{DASHBOARD_JSON}}"
            };
            _services.PromptLibrary.AddOrUpdatePrompt(newPrompt);
            Prompts.Add(newPrompt);
            SelectedPrompt = newPrompt;
        }

        private void RemovePrompt()
        {
            if (SelectedPrompt == null) return;
            _services.PromptLibrary.RemovePrompt(SelectedPrompt.PromptId);
            Prompts.Remove(SelectedPrompt);
            SelectedPrompt = null;
        }

        private void SavePrompt()
        {
            if (SelectedPrompt == null) return;

            // Check if mail body variable is present
            if (!SelectedPrompt.Body.Contains("{{MAIL_BODY}}") && !SelectedPrompt.Body.Contains("{{MAIL_BODY_LATEST}}"))
            {
                var result = MessageBox.Show(
                    Properties.Resources.Current_Prompt_does_not_contai_b912e4, 
                    Properties.Resources.Missing_Email_Body_Variable, 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Warning);
                
                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }

            _services.PromptLibrary.AddOrUpdatePrompt(SelectedPrompt);
        }

        private async void RefreshOllamaModels()
        {
            try
            {
                OllamaModels.Clear();
                var models = await _services.OllamaModels.GetModelsAsync();
                foreach (var model in models)
                {
                    OllamaModels.Add(model);
                }
                
                if (IsOllamaProvider && string.IsNullOrEmpty(LlmConfiguration.ModelName) && OllamaModels.Count > 0)
                {
                    LlmConfiguration.ModelName = OllamaModels[0];
                    RaisePropertyChanged(nameof(LlmConfiguration));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error refreshing Ollama models: {ex.Message}");
            }
        }

        private void SaveLlmConfig(bool showConfirmation = true)
        {
            _services.LlmConfigurations.SaveGlobalConfiguration(LlmConfiguration);
            if (showConfirmation)
            {
                MessageBox.Show(Properties.Resources.LLM_configuration_saved);
            }
        }

        private void AddFile()
        {
            if (SelectedTemplate == null) return;
            if (SelectedTemplate.AttachmentPaths == null) SelectedTemplate.AttachmentPaths = new List<string>();

            var dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == true)
            {
                if (!SelectedTemplate.AttachmentPaths.Contains(dialog.FileName))
                {
                    SelectedTemplate.AttachmentPaths.Add(dialog.FileName);
                    _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                    RaisePropertyChanged(nameof(TemplateFiles));
                }
            }
        }

        private void RemoveFile(string? filePath)
        {
            if (SelectedTemplate == null || filePath == null) return;
            if (SelectedTemplate.AttachmentPaths == null) return;

            if (SelectedTemplate.AttachmentPaths.Remove(filePath))
            {
                _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                RaisePropertyChanged(nameof(TemplateFiles));
            }
        }

        public void AddFiles(IEnumerable<string> files)
        {
            if (SelectedTemplate == null) return;
            if (SelectedTemplate.AttachmentPaths == null) SelectedTemplate.AttachmentPaths = new List<string>();

            bool changed = false;
            foreach (var file in files)
            {
                if (!SelectedTemplate.AttachmentPaths.Contains(file))
                {
                    SelectedTemplate.AttachmentPaths.Add(file);
                    changed = true;
                }
            }
            
            if (changed)
            {
                _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                RaisePropertyChanged(nameof(TemplateFiles));
            }
        }

        private void UpdateSelectableTemplates()
        {
            _allSelectableTemplates.Clear();
            var currentIds = _selectedPrompt?.TemplateOverrideId?.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(id => id.Trim())
                .ToHashSet() ?? new HashSet<string>();

            var seenIds = new HashSet<string>();
            foreach (var template in Templates)
            {
                if (seenIds.Add(template.TemplateId))
                {
                    _allSelectableTemplates.Add(new SelectableTemplateViewModel(
                        template.TemplateId, 
                        template.DisplayName, 
                        currentIds.Contains(template.TemplateId), 
                        OnTemplateSelectionChanged));
                }
            }
            SelectableTemplatesView.Refresh();
            RaisePropertyChanged(nameof(SelectedTemplatesSummary));
            TemplateSearchText = SelectedTemplatesSummary;
        }

        private bool FilterTemplates(object item)
        {
            if (string.IsNullOrWhiteSpace(_templateSearchText)) return true;
            // If the search text matches the summary (user hasn't typed a new search), show all
            if (_templateSearchText == SelectedTemplatesSummary) return true;

            if (item is SelectableTemplateViewModel template)
            {
                return template.DisplayName.IndexOf(_templateSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        private void OnTemplateSelectionChanged()
        {
            if (_selectedPrompt == null) return;
            var selectedIds = _allSelectableTemplates.Where(t => t.IsSelected).Select(t => t.TemplateId);
            _selectedPrompt.TemplateOverrideId = string.Join(",", selectedIds);
            RaisePropertyChanged(nameof(SelectedTemplatesSummary));
            TemplateSearchText = SelectedTemplatesSummary;
        }

        private void SaveAll()
        {
            if (SelectedTemplate != null)
                _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
            if (SelectedPrompt != null)
                _services.PromptLibrary.AddOrUpdatePrompt(SelectedPrompt);
            
            SaveLlmConfig(false);
            MessageBox.Show(Properties.Resources.All_changes_saved, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void AddScript()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Python Scripts (*.py)|*.py",
                Title = "Select Python Script"
            };
            
            if (dialog.ShowDialog() == true)
            {
                var scriptPath = dialog.FileName;
                var fileName = System.IO.Path.GetFileName(scriptPath);
                var destPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts", fileName);
                
                try
                {
                    if (!System.IO.File.Exists(destPath))
                    {
                        System.IO.File.Copy(scriptPath, destPath);
                    }
                    else if (scriptPath != destPath)
                    {
                        if (MessageBox.Show($"Script '{fileName}' already exists. Overwrite?", "Confirm Overwrite", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            System.IO.File.Copy(scriptPath, destPath, true);
                        }
                        else
                        {
                            return;
                        }
                    }
                    
                    RefreshScripts();
                    SelectedScript = Scripts.FirstOrDefault(s => s.ScriptPath == destPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error adding script: {ex.Message}");
                }
            }
        }

        private void RemoveScript()
        {
            if (SelectedScript == null) return;
            if (MessageBox.Show($"Are you sure you want to delete script '{SelectedScript.DisplayName}'?", "Confirm Delete", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try 
                {
                    if (System.IO.File.Exists(SelectedScript.ScriptPath))
                    {
                        System.IO.File.Delete(SelectedScript.ScriptPath);
                    }
                    RefreshScripts();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting script: {ex.Message}");
                }
            }
        }

        public void SaveScript()
        {
            if (SelectedScript == null) return;
            
            SelectedScript.AssociatedTemplateIds = _allSelectableTemplatesForScript
                .Where(t => t.IsSelected)
                .Select(t => t.TemplateId)
                .ToList();
                
            _services.PythonScripts.UpdateScriptMetadata(SelectedScript);
            // Removed MessageBox to avoid interruption on auto-save
        }

        private void RefreshScripts()
        {
            Scripts.Clear();
            foreach (var script in _services.PythonScripts.DiscoverScripts())
            {
                Scripts.Add(script);
            }
        }

        private void UpdateScriptTemplateSelection()
        {
            _allSelectableTemplatesForScript.Clear();
            foreach (var template in Templates)
            {
                var isSelected = SelectedScript?.AssociatedTemplateIds?.Contains(template.TemplateId) ?? false;
                _allSelectableTemplatesForScript.Add(new SelectableTemplateViewModel(
                    template.TemplateId, 
                    template.DisplayName, 
                    isSelected, 
                    () => SaveScript()
                ));
            }
            SelectableTemplatesForScriptView.Refresh();
        }

        private async Task ExportBackupAsync()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "OSEM Backup (*.osembak)|*.osembak",
                FileName = $"OSEM_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.osembak"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    await _services.BackupService.ExportBackupAsync(dialog.FileName);
                    MessageBox.Show(Properties.Resources.Backup_export_successful, Properties.Resources.Export, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(Properties.Resources.Export_failed_ex_Message, ex.Message), Properties.Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task ImportBackupAsync()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "OSEM Backup (*.osembak)|*.osembak"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    await _services.BackupService.ImportBackupAsync(dialog.FileName);
                    MessageBox.Show(Properties.Resources.Backup_import_successful_Pleas_5048f2, Properties.Resources.Import, MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // Refresh lists
                    Templates.Clear();
                    foreach(var t in _services.DashboardTemplates.GetTemplates()) Templates.Add(t);
                    
                    Prompts.Clear();
                    foreach(var p in _services.PromptLibrary.GetPrompts()) Prompts.Add(p);
                    
                    Scripts.Clear();
                    foreach(var s in _services.PythonScripts.DiscoverScripts()) Scripts.Add(s);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(Properties.Resources.Import_failed_ex_Message, ex.Message), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool FilterScriptTemplates(object obj)
        {
            if (string.IsNullOrWhiteSpace(ScriptTemplateSearchText)) return true;
            if (obj is SelectableTemplateViewModel vm)
            {
                return vm.DisplayName.IndexOf(ScriptTemplateSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        private void LoadMonitoredFolders()
        {
            MonitoredFolders.Clear();
            if (Properties.Settings.Default.MonitoredFolders != null)
            {
                var session = _services.OutlookApplication.Session;
                foreach (string entryId in Properties.Settings.Default.MonitoredFolders)
                {
                    try
                    {
                        var folder = session.GetFolderFromID(entryId) as Outlook.Folder;
                        if (folder != null)
                        {
                            MonitoredFolders.Add(new MonitoredFolderViewModel
                            {
                                Name = folder.Name,
                                Path = folder.FolderPath,
                                EntryId = entryId
                            });
                        }
                    }
                    catch
                    {
                        // Folder might be deleted or inaccessible
                        MonitoredFolders.Add(new MonitoredFolderViewModel
                        {
                            Name = "Unknown/Deleted",
                            Path = "Unknown",
                            EntryId = entryId
                        });
                    }
                }
            }
        }

        private void AddFolder()
        {
            try
            {
                var folder = _services.OutlookApplication.Session.PickFolder();
                if (folder != null)
                {
                    if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        var entryId = folder.EntryID;
                        
                        // Check for duplicates
                        if (MonitoredFolders.Any(f => f.EntryId == entryId))
                        {
                            MessageBox.Show(Properties.Resources.This_folder_is_already_in_the_monitor_list, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        MonitoredFolders.Add(new MonitoredFolderViewModel
                        {
                            Name = folder.Name,
                            Path = folder.FolderPath,
                            EntryId = entryId
                        });

                        SaveMonitoredFolders();
                        _services.EventMonitor.RefreshCustomMonitors();
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.Please_select_an_email_folder, Properties.Resources.Info, MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Properties.Resources.Failed_to_add_folder_ex_Message, ex.Message), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveFolder(MonitoredFolderViewModel? folder)
        {
            if (folder != null && MonitoredFolders.Contains(folder))
            {
                MonitoredFolders.Remove(folder);
                SaveMonitoredFolders();
                _services.EventMonitor.RefreshCustomMonitors();
            }
        }

        private void SaveMonitoredFolders()
        {
            if (Properties.Settings.Default.MonitoredFolders == null)
            {
                Properties.Settings.Default.MonitoredFolders = new System.Collections.Specialized.StringCollection();
            }
            
            Properties.Settings.Default.MonitoredFolders.Clear();
            foreach (var folder in MonitoredFolders)
            {
                Properties.Settings.Default.MonitoredFolders.Add(folder.EntryId);
            }
            Properties.Settings.Default.Save();
        }

        public string EventFilesStoragePath
        {
            get => Properties.Settings.Default.EventFilesStoragePath;
            set
            {
                if (Properties.Settings.Default.EventFilesStoragePath != value)
                {
                    Properties.Settings.Default.EventFilesStoragePath = value;
                    Properties.Settings.Default.Save();
                    RaisePropertyChanged();
                }
            }
        }

        private void BrowseStoragePath()
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    EventFilesStoragePath = dialog.SelectedPath;
                }
            }
        }
    }

    public class RegexEntry : ViewModelBase
    {
        private string _regexPattern = string.Empty;
        private string _fieldName = string.Empty;
        public Action<string, string>? OnRename { get; set; }

        public string FieldName 
        { 
            get => _fieldName; 
            set
            {
                if (_fieldName == value) return;
                var oldName = _fieldName;
                _fieldName = value;
                OnRename?.Invoke(oldName, value);
                RaisePropertyChanged();
            }
        }

        public string RegexPattern
        {
            get => _regexPattern;
            set
            {
                if (_regexPattern == value) return;
                _regexPattern = value;
                RaisePropertyChanged();
            }
        }
    }

    public class PromptVariable
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string InsertText { get; set; } = string.Empty;
    }

    public class SelectableTemplateViewModel : ViewModelBase
    {
        private bool _isSelected;
        private readonly Action _onSelectionChanged;

        public SelectableTemplateViewModel(string templateId, string displayName, bool isSelected, Action onSelectionChanged)
        {
            TemplateId = templateId;
            DisplayName = displayName;
            _isSelected = isSelected;
            _onSelectionChanged = onSelectionChanged;
        }

        public string TemplateId { get; }
        public string DisplayName { get; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (SetProperty(ref _isSelected, value))
                {
                    _onSelectionChanged?.Invoke();
                }
            }
        }

        public override string ToString() => DisplayName;
    }
}
