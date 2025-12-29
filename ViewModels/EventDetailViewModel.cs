#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.IO;
using WinForms = System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;
using OSEMAddIn.Services;

namespace OSEMAddIn.ViewModels
{
    internal sealed class EventDetailViewModel : ViewModelBase
    {
        private const string InternetMessageIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
    private static readonly CultureInfo OutlookFilterCulture = CultureInfo.CreateSpecificCulture("en-US");
    private static readonly TimeSpan ManualCatchUpTimeout = TimeSpan.FromSeconds(45);
    private static bool s_conversationFilterSupported = true;

        private readonly ServiceContainer _services;
        private readonly Dispatcher _dispatcher;
        private EventRecord? _currentEvent;
        private string _eventTitle = string.Empty;
        private EventStatus _status = EventStatus.Open;
        private DashboardTemplate? _selectedTemplate;
        private PromptDefinition? _selectedPrompt;
        private PythonScriptDefinition? _selectedScript;
        private bool _isBusy;
        private string _statusMessage = Properties.Resources.Ready_1;
        private double _progressValue;
        private bool _isProgressIndeterminate;
        private PromptDefinition? _previousPrompt;
        private MailItemViewModel? _selectedMail;
        private CancellationTokenSource? _previewCts;
        private bool _suppressPreview;
        private bool _isSettingSelection;
        private Explorer? _monitoredExplorer;

        public EventDetailViewModel(ServiceContainer services)
        {
            _services = services ?? throw new ArgumentNullException(nameof(services));
            _services.EventRepository.EventChanged += OnEventChanged;
            _services.DashboardTemplates.TemplatesChanged += OnTemplatesChanged;
            _services.PromptLibrary.PromptsChanged += OnPromptsChanged;
            _dispatcher = Dispatcher.CurrentDispatcher;

            // Monitor Explorer selection changes to cancel preview if user interacts
            if (_services.OutlookApplication?.ActiveExplorer() is Explorer explorer)
            {
                _monitoredExplorer = explorer;
                _monitoredExplorer.SelectionChange += OnExplorerSelectionChange;
            }

            DashboardItems = new ObservableCollection<DashboardItemViewModel>();
            MailItems = new ObservableCollection<MailItemViewModel>();
            Attachments = new ObservableCollection<AttachmentItemViewModel>();
            Prompts = new ObservableCollection<PromptDefinition>(_services.PromptLibrary.GetPrompts());
            PythonScripts = new ObservableCollection<PythonScriptDefinition>(_services.PythonScripts.DiscoverScripts());

            BackCommand = new AsyncRelayCommand(async _ => 
            {
                await MarkAllAsReadAsync();
                BackRequested?.Invoke(this, EventId);
            });
            NewDashboardCommand = new RelayCommand(_ => ResetDashboard());
            SaveDashboardCommand = new AsyncRelayCommand(_ => SaveDashboardAsync(), _ => _currentEvent is not null);
            ExtractViaRegexCommand = new AsyncRelayCommand(_ => ExtractViaRegexAsync(), _ => _currentEvent is not null && SelectedMail is not null);
            ExtractViaLlmCommand = new AsyncRelayCommand(_ => ExtractViaLlmAsync(), _ => _currentEvent is not null && SelectedPrompt is not null && SelectedTemplate is not null && SelectedMail is not null);
            ExecutePythonScriptCommand = new AsyncRelayCommand(_ => ExecutePythonScriptAsync(), _ => SelectedScript is not null);
            ArchiveCommand = new AsyncRelayCommand(_ => ArchiveAsync(), _ => _currentEvent?.Status == EventStatus.Open);
            ReopenCommand = new AsyncRelayCommand(_ => ReopenAsync(), _ => _currentEvent?.Status == EventStatus.Archived);
            OpenTemplateFileCommand = new RelayCommand(param => OpenTemplateFile(param as string));
            OpenReplyTemplateCommand = new RelayCommand(_ => EmailTemplateRequested?.Invoke(this, new EmailTemplateRequestEventArgs(EmailTemplateType.Reply)));
            OpenComposeTemplateCommand = new RelayCommand(_ => EmailTemplateRequested?.Invoke(this, new EmailTemplateRequestEventArgs(EmailTemplateType.Compose)));
            RefreshMailCommand = new AsyncRelayCommand(_ => RefreshMailAsync(), _ => _currentEvent is not null);
            RemoveMailCommand = new AsyncRelayCommand(param => RemoveMailAsync(param as MailItemViewModel), CanRemoveMail);
            HandleMailDropCommand = new AsyncRelayCommand(_ => HandleMailDropAsync());
            MarkAllAsReadCommand = new AsyncRelayCommand(_ => MarkAllAsReadAsync());
            RemoveTemplateFileCommand = new AsyncRelayCommand(param => RemoveTemplateFileAsync(param));
            SetSubjectSearchRangeCommand = new RelayCommand(param =>
            {
                if (int.TryParse(param?.ToString(), out int days))
                {
                    SubjectSearchRangeDays = days;
                }
            });
            GenerateFolderCommand = new AsyncRelayCommand(_ => GenerateFolderAsync(), _ => _currentEvent is not null);
            UpdateFolderCommand = new AsyncRelayCommand(_ => UpdateFolderAsync(), _ => _currentEvent is not null && !string.IsNullOrEmpty(LocalFolderPath));
            CopyJsonKeysCommand = new RelayCommand(_ => CopyJsonKeys());
        }

        public event EventHandler<string>? BackRequested;
        public event EventHandler? ManagePromptRequested;
        public event EventHandler<EmailTemplateRequestEventArgs>? EmailTemplateRequested;

        public ObservableCollection<DashboardItemViewModel> DashboardItems { get; }
        public ObservableCollection<MailItemViewModel> MailItems { get; }
        public ObservableCollection<AttachmentItemViewModel> Attachments { get; }
        public ObservableCollection<PromptDefinition> Prompts { get; }
        public ObservableCollection<PythonScriptDefinition> PythonScripts { get; }
        public IReadOnlyList<DashboardTemplate> DashboardTemplates => _services.DashboardTemplates.GetTemplates();
        
        private ObservableCollection<TemplateFileViewModel> _templateFiles = new ObservableCollection<TemplateFileViewModel>();
        public ObservableCollection<TemplateFileViewModel> TemplateFiles
        {
            get => _templateFiles;
            set
            {
                if (_templateFiles == value) return;
                _templateFiles = value;
                RaisePropertyChanged();
            }
        }

        public AsyncRelayCommand BackCommand { get; }
        public RelayCommand NewDashboardCommand { get; }
        public AsyncRelayCommand SaveDashboardCommand { get; }
        public AsyncRelayCommand ExtractViaRegexCommand { get; }
        public AsyncRelayCommand ExtractViaLlmCommand { get; }
        public AsyncRelayCommand ExecutePythonScriptCommand { get; }
        public AsyncRelayCommand ArchiveCommand { get; }
        public AsyncRelayCommand ReopenCommand { get; }
        public RelayCommand OpenTemplateFileCommand { get; }
        public RelayCommand OpenReplyTemplateCommand { get; }
        public RelayCommand OpenComposeTemplateCommand { get; }
    public AsyncRelayCommand RefreshMailCommand { get; }
        public AsyncRelayCommand RemoveMailCommand { get; }
        public AsyncRelayCommand HandleMailDropCommand { get; }
        public AsyncRelayCommand MarkAllAsReadCommand { get; }
        public AsyncRelayCommand RemoveTemplateFileCommand { get; }
        public AsyncRelayCommand GenerateFolderCommand { get; }
        public AsyncRelayCommand UpdateFolderCommand { get; }
        public RelayCommand CopyJsonKeysCommand { get; }

        public string LocalFolderPath
        {
            get => _currentEvent?.LocalFolderPath ?? string.Empty;
            set
            {
                if (_currentEvent is null || _currentEvent.LocalFolderPath == value) return;
                _currentEvent.LocalFolderPath = value;
                RaisePropertyChanged();
                UpdateFolderCommand.RaiseCanExecuteChanged();
            }
        }

        private int _subjectSearchRangeDays = 30;
        public int SubjectSearchRangeDays
        {
            get => _subjectSearchRangeDays;
            set
            {
                if (_subjectSearchRangeDays == value) return;
                _subjectSearchRangeDays = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(IsRange30Days));
                RaisePropertyChanged(nameof(IsRange45Days));
                RaisePropertyChanged(nameof(IsRange60Days));
                RaisePropertyChanged(nameof(IsRange180Days));
                RaisePropertyChanged(nameof(IsRangeAll));
            }
        }

        public bool IsRange30Days => SubjectSearchRangeDays == 30;
        public bool IsRange45Days => SubjectSearchRangeDays == 45;
        public bool IsRange60Days => SubjectSearchRangeDays == 60;
        public bool IsRange180Days => SubjectSearchRangeDays == 180;
        public bool IsRangeAll => SubjectSearchRangeDays == 3650;

        public RelayCommand SetSubjectSearchRangeCommand { get; }

        public string EventId => _currentEvent?.EventId ?? string.Empty;

        public string EventTitle
        {
            get => _eventTitle;
            set
            {
                if (_eventTitle == value)
                {
                    return;
                }

                _eventTitle = value ?? string.Empty;
                RaisePropertyChanged();
                _ = PersistTitleAsync();
            }
        }

        public EventStatus Status
        {
            get => _status;
            private set
            {
                if (_status == value)
                {
                    return;
                }
                _status = value;
                RaisePropertyChanged();
                ArchiveCommand.RaiseCanExecuteChanged();
                ReopenCommand.RaiseCanExecuteChanged();
            }
        }

        public DashboardTemplate? SelectedTemplate
        {
            get => _selectedTemplate;
            set
            {
                if (_selectedTemplate == value)
                {
                    return;
                }

                _selectedTemplate = value;
                RaisePropertyChanged();
                UpdateTemplateFiles();
                UpdatePrompts();
                if (value is not null)
                {
                    EnsureDashboardFields(value);
                    _ = PersistTemplateAsync();
                }
            }
        }

        private string _displayColumnSource = "Custom";
        public string DisplayColumnSource
        {
            get => _displayColumnSource;
            set
            {
                if (_displayColumnSource == value) return;
                _displayColumnSource = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(IsCustomDisplayColumn));
                _ = PersistDisplayColumnAsync();
            }
        }

        private string _displayColumnCustomValue = string.Empty;
        public string DisplayColumnCustomValue
        {
            get => _displayColumnCustomValue;
            set
            {
                if (_displayColumnCustomValue == value) return;
                _displayColumnCustomValue = value;
                RaisePropertyChanged();
                _ = PersistDisplayColumnAsync();
            }
        }

        public bool IsCustomDisplayColumn => string.Equals(DisplayColumnSource, "Custom", StringComparison.OrdinalIgnoreCase);

        public ObservableCollection<string> AvailableDisplaySources { get; } = new();

        public PromptDefinition? SelectedPrompt
        {
            get => _selectedPrompt;
            set
            {
                if (_selectedPrompt == value)
                {
                    return;
                }

                _selectedPrompt = value;
                RaisePropertyChanged();
                if (_selectedPrompt is not null && string.Equals(_selectedPrompt.PromptId, "prompt.manage", StringComparison.Ordinal))
                {
                    ManagePromptRequested?.Invoke(this, EventArgs.Empty);
                    _selectedPrompt = _previousPrompt;
                    RaisePropertyChanged();
                }
                else
                {
                    _previousPrompt = _selectedPrompt;
                }

                ExtractViaLlmCommand.RaiseCanExecuteChanged();
            }
        }

        public PythonScriptDefinition? SelectedScript
        {
            get => _selectedScript;
            set
            {
                if (_selectedScript == value)
                {
                    return;
                }

                _selectedScript = value;
                RaisePropertyChanged();
                ExecutePythonScriptCommand.RaiseCanExecuteChanged();
            }
        }

        public MailItemViewModel? SelectedMail
        {
            get => _selectedMail;
            set
            {
                if (_selectedMail == value)
                {
                    return;
                }

                _selectedMail = value;
                RaisePropertyChanged();
                RemoveMailCommand.RaiseCanExecuteChanged();
                ExtractViaLlmCommand.RaiseCanExecuteChanged();
                ExtractViaRegexCommand.RaiseCanExecuteChanged();
                
                if (!_suppressPreview)
                {
                    _previewCts?.Cancel();
                    _previewCts = new CancellationTokenSource();
                    _ = PreviewMailAsync(value, _previewCts.Token);
                }
            }
        }

        private async Task PreviewMailAsync(MailItemViewModel? mail, CancellationToken token)
        {
            if (mail is null) return;

            try
            {
                // Debounce to allow double-click to happen first
                await Task.Delay(300, token);
            }
            catch (TaskCanceledException)
            {
                return;
            }

            if (token.IsCancellationRequested) return;

            await _dispatcher.InvokeAsync(async () =>
            {
                if (token.IsCancellationRequested) return;

                try
                {
                    var app = _services.OutlookApplication;
                    if (app is null) return;

                    var explorer = app.ActiveExplorer();
                    if (explorer is null) return;

                    // Update monitored explorer if changed
                    if (_monitoredExplorer != explorer)
                    {
                        if (_monitoredExplorer != null)
                            _monitoredExplorer.SelectionChange -= OnExplorerSelectionChange;
                        
                        _monitoredExplorer = explorer;
                        _monitoredExplorer.SelectionChange += OnExplorerSelectionChange;
                    }

                    // 1. Resolve item first to check parent
                    var item = ResolveMail(app, mail);
                    if (item is null) return;

                    try
                    {
                        // If already selected, do nothing
                        if (explorer.Selection.Count == 1 && explorer.Selection[1] is MailItem selected && selected.EntryID == item.EntryID)
                        {
                            Marshal.ReleaseComObject(selected);
                            return;
                        }

                        _isSettingSelection = true;

                        // 2. Switch folder if necessary
                        bool folderSwitched = false;
                        if (item.Parent is MAPIFolder parent)
                        {
                            try
                            {
                                if (explorer.CurrentFolder.EntryID != parent.EntryID)
                                {
                                    explorer.CurrentFolder = parent;
                                    folderSwitched = true;
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(parent);
                            }
                        }

                        if (folderSwitched)
                        {
                            // Release the old item reference as it might be unstable during view switch
                            Marshal.ReleaseComObject(item);
                            item = null;

                            // Wait for view to load
                            try 
                            {
                                await Task.Delay(500, token);
                            }
                            catch (TaskCanceledException) { return; }

                            if (token.IsCancellationRequested) return;

                            // Re-resolve the item in the new context
                            item = ResolveMail(app, mail);
                            if (item is null) return;
                        }

                        // 3. Select
                        // Ensure we are in the right folder
                        if (explorer.CurrentFolder.EntryID == ((MAPIFolder)item.Parent).EntryID)
                        {
                            // Activate the explorer window to ensure selection is processed visually
                            explorer.Activate();
                            
                            explorer.ClearSelection();
                            
                            try 
                            {
                                explorer.AddToSelection(item);
                            }
                            catch
                            {
                                // Retry once if first attempt fails (e.g. view not ready)
                                try 
                                {
                                    await Task.Delay(500, token);
                                    if (!token.IsCancellationRequested)
                                        explorer.AddToSelection(item);
                                }
                                catch { /* Ignore second failure or cancellation */ }
                            }

                            // Scroll to selection if possible (Outlook 2010+)
                            // This ensures the item is visible in the list
                            if (explorer.Selection.Count > 0)
                            {
                                // No direct ScrollToSelection, but AddToSelection usually scrolls.
                                // We can try to ensure it's visible by accessing it?
                            }
                        }
                    }
                    finally
                    {
                        _isSettingSelection = false;
                        if (item != null) Marshal.ReleaseComObject(item);
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLogger.Log($"PreviewMailAsync failed: {ex.Message}");
                    _isSettingSelection = false;
                }
            });
        }

        internal ServiceContainer Services => _services;

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (_isBusy == value)
                {
                    return;
                }

                _isBusy = value;
                InvokeOnUiThread(() => RaisePropertyChanged(nameof(IsBusy)));
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage == value) return;
                _statusMessage = value;
                InvokeOnUiThread(() => RaisePropertyChanged(nameof(StatusMessage)));
            }
        }

        public double ProgressValue
        {
            get => _progressValue;
            set
            {
                if (Math.Abs(_progressValue - value) < 0.01) return;
                _progressValue = value;
                InvokeOnUiThread(() => RaisePropertyChanged(nameof(ProgressValue)));
            }
        }

        public bool IsProgressIndeterminate
        {
            get => _isProgressIndeterminate;
            set
            {
                if (_isProgressIndeterminate == value) return;
                _isProgressIndeterminate = value;
                InvokeOnUiThread(() => RaisePropertyChanged(nameof(IsProgressIndeterminate)));
            }
        }

        public async Task LoadAsync(string eventId)
        {
            DebugLogger.Log($"EventDetailViewModel: Loading event ID: {eventId}");
            
            await _dispatcher.InvokeAsync(Clear);

            if (string.IsNullOrWhiteSpace(eventId))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(eventId));
            }

            var record = await _services.EventRepository.GetByIdAsync(eventId);
            if (record is null)
            {
                DebugLogger.Log($"EventDetailViewModel: Record not found for ID: {eventId}");
                await _dispatcher.InvokeAsync(() => 
                {
                    EventTitle = "Event not found";
                });
                return;
            }
            DebugLogger.Log($"EventDetailViewModel: Loaded record '{record.EventTitle}' for ID: {eventId}");

            await _dispatcher.InvokeAsync(() =>
            {
                _currentEvent = record;
                EventTitle = record.EventTitle;
                Status = record.Status;
                
                PopulateDashboard(record);
                
                var template = _services.DashboardTemplates.FindById(record.DashboardTemplateId) ?? _services.DashboardTemplates.GetTemplates().FirstOrDefault();
                _selectedTemplate = template;
                RaisePropertyChanged(nameof(SelectedTemplate));
                UpdatePrompts();
                
                if (template != null)
                {
                    EnsureDashboardFields(template);
                }

                PopulateMail(record);
                PopulateAttachments(record);

                // Set these last to ensure AvailableDisplaySources is populated
                DisplayColumnSource = record.DisplayColumnSource;
                DisplayColumnCustomValue = record.DisplayColumnCustomValue;
            });

            // Fix: Do not automatically mark as read on load. 
            // User wants to see the highlight to confirm new mail.
            // They can manually mark as read or we can do it on specific actions.
            /*
            if (record.Emails.Any(e => e.IsNewOrUpdated))
            {
                foreach (var email in record.Emails)
                {
                    email.IsNewOrUpdated = false;
                }

                await _services.EventRepository.UpdateAsync(record);
            }
            */

            await ValidateMailEntriesAsync(record);
            RefreshMailCommand.RaiseCanExecuteChanged();
        }

        private void Clear()
        {
            _currentEvent = null;
            EventTitle = "Loading...";
            // Don't reset Status, keep it as is or default
            DashboardItems.Clear();
            MailItems.Clear();
            Attachments.Clear();
            SelectedMail = null;
        }

        private void PopulateDashboard(EventRecord record)
        {
            InvokeOnUiThread(() =>
            {
                // Sync SelectedTemplate from record if needed
                if (!string.IsNullOrEmpty(record.DashboardTemplateId))
                {
                    var template = _services.DashboardTemplates.GetTemplates()
                        .FirstOrDefault(t => t.TemplateId == record.DashboardTemplateId);
                    
                    // Update backing field directly to avoid triggering property setter logic (persistence/ensure fields)
                    // because we are loading state, not changing it
                    if (template != null && _selectedTemplate != template)
                    {
                        _selectedTemplate = template;
                        RaisePropertyChanged(nameof(SelectedTemplate));
                    }
                }
                
                // Always update files to include event-specific files
                UpdateTemplateFiles();

                // Capture current values from record
                var recordValues = record.DashboardItems.ToDictionary(i => i.Key, i => i.Value, StringComparer.OrdinalIgnoreCase);
                
                // Determine which fields to show
                IEnumerable<string> fieldsToShow;
                if (SelectedTemplate != null && SelectedTemplate.Fields != null)
                {
                    // If a template is selected, show its fields
                    fieldsToShow = SelectedTemplate.Fields;
                }
                else
                {
                    // Otherwise show whatever is in the record
                    fieldsToShow = recordValues.Keys;
                }

                DashboardItems.Clear();
                foreach (var key in fieldsToShow)
                {
                    var value = recordValues.TryGetValue(key, out var v) ? v : string.Empty;
                    DashboardItems.Add(new DashboardItemViewModel(key, value, OnDashboardItemChanged));
                }
                
                UpdateAvailableDisplaySources();
            });
        }

        private void OnDashboardItemChanged()
        {
            if (_currentEvent is null) return;
            
            _currentEvent.DashboardItems = DashboardItems.Select(vm => vm.ToModel()).ToList();
            _ = _services.EventRepository.UpdateAsync(_currentEvent);
        }

        private void PopulateMail(EventRecord record)
        {
            InvokeOnUiThread(() =>
            {
                DebugLogger.Log($"PopulateMail: Updating MailItems for Event='{record.EventId}'. Count={record.Emails.Count}");
                var previous = SelectedMail;
                MailItems.Clear();
                foreach (var mail in record.Emails.Where(e => !e.IsRemoved).OrderByDescending(e => e.ReceivedOn))
                {
                    MailItems.Add(new MailItemViewModel(mail));
                }
                DebugLogger.Log($"PopulateMail: Updated MailItems. New Count={MailItems.Count}");

                if (previous is not null)
                {
                    _suppressPreview = true;
                    try
                    {
                        SelectedMail = MailItems.FirstOrDefault(m =>
                            string.Equals(m.EntryId, previous.EntryId, StringComparison.OrdinalIgnoreCase) ||
                            (!string.IsNullOrEmpty(previous.InternetMessageId) && string.Equals(m.InternetMessageId, previous.InternetMessageId, StringComparison.OrdinalIgnoreCase)));
                    }
                    finally
                    {
                        _suppressPreview = false;
                    }
                }
            });
        }

        private void PopulateAttachments(EventRecord record)
        {
            InvokeOnUiThread(() =>
            {
                Attachments.Clear();
                // Handle potential duplicate EntryIds safely
                var emailLookup = record.Emails
                    .Where(e => !string.IsNullOrEmpty(e.EntryId))
                    .GroupBy(e => e.EntryId, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                var viewModels = new List<AttachmentItemViewModel>();
                foreach (var attachment in record.Attachments)
                {
                    EmailItem? sourceEmail = null;
                    if (!string.IsNullOrEmpty(attachment.SourceMailEntryId))
                    {
                        emailLookup.TryGetValue(attachment.SourceMailEntryId, out sourceEmail);
                    }
                    viewModels.Add(new AttachmentItemViewModel(attachment, sourceEmail));
                }

                foreach (var vm in viewModels.OrderBy(v => v.SortPriority).ThenByDescending(v => v.ReceivedOn))
                {
                    Attachments.Add(vm);
                }
            });
        }

        private void EnsureDashboardFields(DashboardTemplate template)
        {
            InvokeOnUiThread(() =>
            {
                var existingKeys = DashboardItems.Select(item => item.Key).ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var field in template.Fields)
                {
                    if (!existingKeys.Contains(field))
                    {
                        DashboardItems.Add(new DashboardItemViewModel(field, string.Empty, OnDashboardItemChanged));
                    }
                }

                var fieldLookup = template.Fields.ToHashSet(StringComparer.OrdinalIgnoreCase);
                for (var i = DashboardItems.Count - 1; i >= 0; i--)
                {
                    if (!fieldLookup.Contains(DashboardItems[i].Key))
                    {
                        DashboardItems.RemoveAt(i);
                    }
                }
                UpdateAvailableDisplaySources();
            });
        }

        private async Task SaveDashboardAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            IsBusy = true;
            _services.BusyState.SetBusy("Saving dashboard...");
            try
            {
                _currentEvent.EventTitle = EventTitle;
                _currentEvent.Status = Status;
                _currentEvent.DashboardTemplateId = SelectedTemplate?.TemplateId ?? string.Empty;
                _currentEvent.DashboardItems = DashboardItems.Select(item => item.ToModel()).ToList();
                await _services.EventRepository.UpdateAsync(_currentEvent);
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
            }
        }

        private async Task ExtractViaRegexAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            if (SelectedMail is null)
            {
                return;
            }

            var application = _services.OutlookApplication;
            if (application is null)
            {
                return;
            }

            MailItem? mailItem = null;
            IsBusy = true;
            _services.BusyState.SetBusy("Extracting via Regex...");
            try
            {
                mailItem = ResolveMail(application, SelectedMail);
                if (mailItem is null)
                {
                    return;
                }

                if (SelectedTemplate is null)
                {
                    return;
                }

                var results = _services.RegexExtraction.Extract(mailItem, SelectedTemplate);
                ApplyExtractionResults(results);
                await SaveDashboardAsync();
            }
            finally
            {
                ReleaseComObject(mailItem);
                IsBusy = false;
                _services.BusyState.ClearBusy();
            }
        }

        private async Task ExtractViaLlmAsync()
        {
            if (_currentEvent is null || SelectedPrompt is null || SelectedTemplate is null || SelectedMail is null)
            {
                return;
            }

            var application = _services.OutlookApplication;
            if (application is null)
            {
                return;
            }

            MailItem? mailItem = null;
            IsBusy = true;
            _services.BusyState.SetBusy(Properties.Resources.Requesting_LLM_extraction);
            StatusMessage = Properties.Resources.Requesting_LLM_extraction;
            IsProgressIndeterminate = true;

            try
            {
                mailItem = ResolveMail(application, SelectedMail);
                if (mailItem is null)
                {
                    StatusMessage = Properties.Resources.Unable_to_parse_selected_email;
                    IsBusy = false;
                    _services.BusyState.ClearBusy();
                    return;
                }

                var configuration = _services.LlmConfigurations.GetEffectiveConfiguration(SelectedTemplate.TemplateId);
                var results = await _services.LlmExtraction.ExtractAsync(SelectedPrompt, SelectedTemplate, mailItem, configuration);
                ApplyExtractionResults(results);
                await SaveDashboardAsync();
                StatusMessage = Properties.Resources.Extraction_completed;
            }
            catch (System.Exception ex)
            {
                StatusMessage = string.Format(Properties.Resources.Extraction_failed_ex_Message, ex.Message);
                DebugLogger.Log($"LLM Extraction failed: {ex}");
            }
            finally
            {
                ReleaseComObject(mailItem);
                IsBusy = false;
                _services.BusyState.ClearBusy();
                IsProgressIndeterminate = false;
            }
        }

        private void ApplyExtractionResults(Dictionary<string, string> results)
        {
            InvokeOnUiThread(() =>
            {
                foreach (var item in DashboardItems)
                {
                    if (results.TryGetValue(item.Key, out var value))
                    {
                        // If the new value is empty, and we already have a value, don't overwrite it.
                        if (string.IsNullOrWhiteSpace(value) && !string.IsNullOrWhiteSpace(item.Value))
                        {
                            continue;
                        }
                        item.Value = value;
                    }
                }
            });
        }

        private async Task ExecutePythonScriptAsync()
        {
            if (SelectedScript is null)
            {
                return;
            }

            if (_currentEvent is null)
            {
                return;
            }

            var context = new PythonScriptExecutionContext
            {
                EventId = _currentEvent.EventId,
                EventTitle = _currentEvent.EventTitle,
                DashboardTemplateId = SelectedTemplate?.TemplateId,
                DashboardValues = DashboardItems.ToDictionary(item => item.Key, item => item.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase),
                Emails = _currentEvent.Emails
                    .Where(e => !e.IsRemoved)
                    .Select(e => new EmailItem
                    {
                        EntryId = e.EntryId,
                        StoreId = e.StoreId,
                        ConversationId = e.ConversationId,
                        InternetMessageId = e.InternetMessageId,
                        Sender = e.Sender,
                        To = e.To,
                        Subject = e.Subject,
                        ReceivedOn = e.ReceivedOn,
                        IsNewOrUpdated = e.IsNewOrUpdated,
                        IsRemoved = e.IsRemoved
                    })
                    .ToList(),
                Attachments = _currentEvent.Attachments.Select(a => new AttachmentItem
                {
                    Id = a.Id,
                    FileName = a.FileName,
                    FileType = a.FileType,
                    FileSizeBytes = a.FileSizeBytes,
                    SourceMailEntryId = a.SourceMailEntryId
                }).ToList()
            };

            IsBusy = true;
            _services.BusyState.SetBusy($"Executing script: {SelectedScript.DisplayName}...");
            try
            {
                await _services.PythonScripts.ExecuteAsync(SelectedScript, context);
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
            }
        }

        private async Task ArchiveAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            await _services.EventRepository.ArchiveAsync(new[] { _currentEvent.EventId });
        }
        private async Task ReopenAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            DebugLogger.Log($"ReopenAsync called for EventId={_currentEvent.EventId}, Status={_currentEvent.Status}");

            if (_currentEvent.Status != EventStatus.Archived)
            {
                System.Windows.MessageBox.Show($"Cannot reopen event. Current status is {_currentEvent.Status}. Expected: Archived.");
                return;
            }

            var message = "Are you sure you want to reopen this event?";
            if (System.Windows.MessageBox.Show(message, "Confirm Reopen", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question) != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }

            IsBusy = true;
            _services.BusyState.SetBusy("Reopening event...");
            try
            {
                await _services.EventRepository.ReopenAsync(_currentEvent.EventId);
                DebugLogger.Log("ReopenAsync completed successfully.");
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"ReopenAsync failed: {ex}");
                System.Windows.MessageBox.Show($"Reopen failed: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
            }
        }

        private async Task PersistTitleAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            if (string.Equals(_currentEvent.EventTitle, EventTitle, StringComparison.Ordinal))
            {
                return;
            }

            _currentEvent.EventTitle = EventTitle;
            await _services.EventRepository.UpdateAsync(_currentEvent);
        }

        private async Task PersistTemplateAsync()
        {
            if (_currentEvent is null || SelectedTemplate is null) return;
            if (string.Equals(_currentEvent.DashboardTemplateId, SelectedTemplate.TemplateId, StringComparison.Ordinal)) return;

            _currentEvent.DashboardTemplateId = SelectedTemplate.TemplateId;
            await _services.EventRepository.UpdateAsync(_currentEvent);
        }

        private async Task PersistDisplayColumnAsync()
        {
            if (_currentEvent is null) return;
            
            bool changed = false;
            if (!string.Equals(_currentEvent.DisplayColumnSource, DisplayColumnSource, StringComparison.Ordinal))
            {
                _currentEvent.DisplayColumnSource = DisplayColumnSource;
                changed = true;
            }
            if (!string.Equals(_currentEvent.DisplayColumnCustomValue, DisplayColumnCustomValue, StringComparison.Ordinal))
            {
                _currentEvent.DisplayColumnCustomValue = DisplayColumnCustomValue;
                changed = true;
            }

            if (changed)
            {
                await _services.EventRepository.UpdateAsync(_currentEvent);
            }
        }

        private void UpdateAvailableDisplaySources()
        {
            var newSources = DashboardItems.Select(item => item.Key).ToList();
            
            // Remove items that are no longer present
            for (int i = AvailableDisplaySources.Count - 1; i >= 0; i--)
            {
                if (!newSources.Contains(AvailableDisplaySources[i]))
                {
                    AvailableDisplaySources.RemoveAt(i);
                }
            }

            // Add new items
            foreach (var source in newSources)
            {
                if (!AvailableDisplaySources.Contains(source))
                {
                    AvailableDisplaySources.Add(source);
                }
            }
        }

        private void ResetDashboard()
        {
            InvokeOnUiThread(() =>
            {
                DashboardItems.Clear();
                if (SelectedTemplate is not null)
                {
                    EnsureDashboardFields(SelectedTemplate);
                }
            });
        }

        private void OnEventChanged(object? sender, EventChangedEventArgs e)
        {
            if (_currentEvent is null || !string.Equals(e.Record.EventId, _currentEvent.EventId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            DebugLogger.Log($"EventDetailViewModel.OnEventChanged: Received update for Event='{e.Record.EventId}' Reason='{e.ChangeReason}'");

            InvokeOnUiThread(() =>
            {
                _currentEvent = e.Record;
                Status = e.Record.Status;
                EventTitle = e.Record.EventTitle;

                PopulateDashboard(e.Record);
                PopulateMail(e.Record);
                PopulateAttachments(e.Record);

                // Always update the backing field and raise property changed to ensure UI sync,
                // especially since ComboBox text might be cleared when ItemsSource (AvailableDisplaySources) is updated.
                _displayColumnSource = e.Record.DisplayColumnSource;
                RaisePropertyChanged(nameof(DisplayColumnSource));
                RaisePropertyChanged(nameof(IsCustomDisplayColumn));

                _displayColumnCustomValue = e.Record.DisplayColumnCustomValue;
                RaisePropertyChanged(nameof(DisplayColumnCustomValue));

                RefreshMailCommand.RaiseCanExecuteChanged();
                RemoveMailCommand.RaiseCanExecuteChanged();
            });
        }

        private void OnTemplatesChanged(object? sender, EventArgs e)
        {
            _dispatcher.Invoke(() =>
            {
                var currentId = SelectedTemplate?.TemplateId;
                RaisePropertyChanged(nameof(DashboardTemplates));
                
                if (currentId != null)
                {
                    // Re-sync SelectedTemplate to the latest instance from the service
                    var match = _services.DashboardTemplates.FindById(currentId);
                    if (match != null && !ReferenceEquals(match, SelectedTemplate))
                    {
                        SelectedTemplate = match;
                    }
                    else
                    {
                        // Even if it's the same instance or null, force an update of the files
                        UpdateTemplateFiles();
                    }
                }
                else
                {
                    UpdateTemplateFiles();
                }
            });
        }

        private void OnPromptsChanged(object? sender, EventArgs e)
        {
            _dispatcher.Invoke(() =>
            {
                UpdatePrompts();
            });
        }

        private void UpdatePrompts()
        {
            var currentSelectionId = SelectedPrompt?.PromptId;
            var allPrompts = _services.PromptLibrary.GetPrompts();

            var filteredPrompts = allPrompts.Where(p =>
                string.IsNullOrEmpty(p.TemplateOverrideId) ||
                (SelectedTemplate != null && p.TemplateOverrideId != null && 
                 p.TemplateOverrideId.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Any(id => string.Equals(id.Trim(), SelectedTemplate.TemplateId, StringComparison.OrdinalIgnoreCase)))
            ).ToList();

            Prompts.Clear();
            foreach (var prompt in filteredPrompts)
            {
                Prompts.Add(prompt);
            }

            if (currentSelectionId != null)
            {
                var newSelection = Prompts.FirstOrDefault(p => p.PromptId == currentSelectionId);
                if (newSelection != null)
                {
                    SelectedPrompt = newSelection;
                }
                else if (Prompts.Count > 0)
                {
                    SelectedPrompt = Prompts[0];
                }
                else
                {
                    SelectedPrompt = null;
                }
            }
            else if (Prompts.Count > 0)
            {
                SelectedPrompt = Prompts[0];
            }
        }

        public void ApplyEmailTemplate(EmailTemplate template)
        {
            if (template is null)
            {
                return;
            }

            var application = _services.OutlookApplication;
            if (application is null)
            {
                return;
            }

            MailItem? sourceMail = null;
            try
            {
                if (SelectedMail is not null)
                {
                    sourceMail = ResolveMail(application, SelectedMail);
                }

                if (template.TemplateType == EmailTemplateType.Reply)
                {
                    if (sourceMail is null && MailItems.FirstOrDefault() is MailItemViewModel fallback)
                    {
                        sourceMail = ResolveMail(application, fallback);
                    }

                    var reply = sourceMail?.Reply();
                    if (reply is null)
                    {
                        return;
                    }

                    reply.Subject = MergeTemplateTokens(template.Subject, sourceMail);
                    reply.HTMLBody = MergeTemplateTokens(template.Body, sourceMail) + reply.HTMLBody;
                    reply.Display();
                }
                else
                {
                    var compose = application.CreateItem(OlItemType.olMailItem) as MailItem;
                    if (compose is null)
                    {
                        return;
                    }

                    compose.Subject = MergeTemplateTokens(template.Subject, sourceMail);
                    compose.HTMLBody = MergeTemplateTokens(template.Body, sourceMail);
                    compose.Display();
                }
            }
            finally
            {
                ReleaseComObject(sourceMail);
            }
        }

        private async Task RemoveMailAsync(MailItemViewModel? mail)
        {
            if (mail is null || _currentEvent is null)
            {
                return;
            }

            await _services.EventRepository.RemoveMailAsync(_currentEvent.EventId, mail.EntryId, mail.InternetMessageId);
            SelectedMail = null;
        }

        private bool CanRemoveMail(object? parameter)
        {
            return _currentEvent is not null && parameter is MailItemViewModel;
        }

        private async Task HandleMailDropAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            var explorer = _services.OutlookApplication.ActiveExplorer();
            if (explorer?.Selection is null || explorer.Selection.Count == 0)
            {
                return;
            }

            for (var index = 1; index <= explorer.Selection.Count; index++)
            {
                if (explorer.Selection[index] is MailItem mail)
                {
                    try
                    {
                        await _services.EventRepository.AddMailToEventAsync(_currentEvent.EventId, mail);
                    }
                    finally
                    {
                        ReleaseComObject(mail);
                    }
                }
            }
        }

        private string MergeTemplateTokens(string input, MailItem? sourceMail)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            var tokens = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["{{EventId}}"] = EventId,
                ["{{EventTitle}}"] = EventTitle,
                ["{{MailSubject}}"] = sourceMail?.Subject ?? string.Empty,
                ["{{MailSender}}"] = sourceMail?.SenderName ?? string.Empty
            };

            var result = input;
            foreach (var kvp in tokens)
            {
                result = result.Replace(kvp.Key, kvp.Value ?? string.Empty);
            }

            return result;
        }

        private MailItem? ResolveMail(Application application, EmailItem? email)
        {
            if (email is null)
            {
                return null;
            }

            return ResolveMail(application, email.EntryId, email.StoreId, email.InternetMessageId);
        }

        private MailItem? ResolveMail(Application application, MailItemViewModel? mail)
        {
            if (mail is null)
            {
                return null;
            }

            return ResolveMail(application, mail.EntryId, mail.StoreId, mail.InternetMessageId);
        }

        private MailItem? ResolveMail(Application application, string? entryId, string? storeId, string? internetMessageId)
        {
            if (application is null || string.IsNullOrEmpty(entryId))
            {
                return null;
            }

            MailItem? item = null;
            var session = application.Session;

            if (!string.IsNullOrEmpty(storeId))
            {
                try
                {
                    item = session.GetItemFromID(entryId, storeId) as MailItem;
                }
                catch (COMException)
                {
                    item = null;
                }
                catch
                {
                    item = null;
                }
            }

            if (item is null)
            {
                try
                {
                    item = session.GetItemFromID(entryId) as MailItem;
                }
                catch (COMException)
                {
                    item = null;
                }
                catch
                {
                    item = null;
                }
            }

            if (item is null && !string.IsNullOrWhiteSpace(internetMessageId))
            {
                item = FindMailByInternetMessageId(application, internetMessageId!);
            }

            return item;
        }

        private MailItem? FindMailByInternetMessageId(Application application, string messageId)
        {
            if (application is null || string.IsNullOrWhiteSpace(messageId))
            {
                return null;
            }

            Stores? stores = null;

            try
            {
                stores = application.Session?.Stores;
            }
            catch (COMException)
            {
                stores = null;
            }
            catch
            {
                stores = null;
            }

            var storeCollection = stores;
            if (storeCollection is null)
            {
                return null;
            }

            try
            {
                for (var index = 1; index <= storeCollection.Count; index++)
                {
                    Store? store = null;
                    try
                    {
                        store = storeCollection[index];
                        if (store is null)
                        {
                            continue;
                        }

                        var candidate = SearchStoreByInternetMessageId(store, messageId);
                        if (candidate is not null)
                        {
                            return candidate;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(store);
                    }
                }
            }
            finally
            {
                ReleaseComObject(storeCollection);
            }

            return null;
        }

        private MailItem? SearchStoreByInternetMessageId(Store store, string messageId)
        {
            if (store is null)
            {
                return null;
            }

            var cutoffUtc = DateTime.UtcNow.AddDays(-7);

            MAPIFolder? inbox = null;
            Folders? inboxChildren = null;

            try
            {
                inbox = SafeGetDefaultFolder(store, OlDefaultFolders.olFolderInbox);
                if (inbox is not null)
                {
                    var inboxMatch = SearchFolder(inbox, messageId, cutoffUtc);
                    if (inboxMatch is not null)
                    {
                        return inboxMatch;
                    }

                    try
                    {
                        inboxChildren = inbox.Folders;
                        if (inboxChildren is not null)
                        {
                            for (var i = 1; i <= inboxChildren.Count; i++)
                            {
                                MAPIFolder? child = null;
                                try
                                {
                                    child = inboxChildren[i];
                                    if (child is null)
                                    {
                                        continue;
                                    }

                                    var childMatch = SearchFolder(child, messageId, cutoffUtc);
                                    if (childMatch is not null)
                                    {
                                        return childMatch;
                                    }
                                }
                                finally
                                {
                                    ReleaseComObject(child);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(inboxChildren);
                        inboxChildren = null;
                    }
                }

                if (SafeGetDefaultFolder(store, OlDefaultFolders.olFolderSentMail) is MAPIFolder sent)
                {
                    var sentMatch = SearchFolder(sent, messageId, cutoffUtc);
                    ReleaseComObject(sent);
                    if (sentMatch is not null)
                    {
                        return sentMatch;
                    }
                }

                if (SafeGetDefaultFolder(store, OlDefaultFolders.olFolderDeletedItems) is MAPIFolder deleted)
                {
                    var deletedMatch = SearchFolder(deleted, messageId, cutoffUtc);
                    ReleaseComObject(deleted);
                    if (deletedMatch is not null)
                    {
                        return deletedMatch;
                    }
                }
            }
            finally
            {
                ReleaseComObject(inboxChildren);
                ReleaseComObject(inbox);
            }

            return null;
        }

        private MailItem? SearchFolder(MAPIFolder folder, string messageId, DateTime cutoffUtc)
        {
            if (folder is null)
            {
                return null;
            }

            Items? items = null;
            Items? filtered = null;
            object? rawItem = null;
            MailItem? result = null;

            try
            {
                items = folder.Items;
                items.Sort("[ReceivedTime]", true);

                var filter = BuildReceivedSinceFilter(cutoffUtc);
                filtered = items.Restrict(filter);

                rawItem = filtered.GetFirst();
                while (rawItem is not null)
                {
                    if (rawItem is MailItem mail)
                    {
                        PropertyAccessor? accessor = null;
                        try
                        {
                            accessor = mail.PropertyAccessor;
                            var candidateId = accessor?.GetProperty(InternetMessageIdProperty) as string;
                            if (!string.IsNullOrEmpty(candidateId) && string.Equals(candidateId, messageId, StringComparison.OrdinalIgnoreCase))
                            {
                                result = mail;
                                break;
                            }
                        }
                        catch (COMException)
                        {
                            // Ignore and continue scanning other items.
                        }
                        finally
                        {
                            if (accessor is not null)
                            {
                                Marshal.ReleaseComObject(accessor);
                            }
                        }

                        Marshal.ReleaseComObject(mail);
                    }
                    else
                    {
                        Marshal.ReleaseComObject(rawItem);
                    }

                    rawItem = filtered.GetNext();
                }
            }
            catch (COMException)
            {
                result = null;
            }
            finally
            {
                if (result is null && rawItem is not null)
                {
                    Marshal.ReleaseComObject(rawItem);
                }

                ReleaseComObject(filtered);
                ReleaseComObject(items);
            }

            return result;
        }

        private static MAPIFolder? SafeGetDefaultFolder(Store store, OlDefaultFolders folderType)
        {
            try
            {
                return store.GetDefaultFolder(folderType);
            }
            catch (COMException)
            {
                return null;
            }
            catch
            {
                return null;
            }
        }

        private async Task RefreshMailAsync()
        {
            if (_currentEvent is null)
            {
                return;
            }

            if (_services.OutlookApplication is null)
            {
                DebugLogger.Log("RefreshMailAsync skipped because Outlook application is unavailable.");
                return;
            }

            IsBusy = true;
            _services.BusyState.SetBusy(Properties.Resources.Preparing_to_refresh);
            IsProgressIndeterminate = false;
            ProgressValue = 0;
            StatusMessage = Properties.Resources.Preparing_to_refresh;
            var eventId = _currentEvent.EventId;

            try
            {
                DebugLogger.Log($"RefreshMailAsync started for Event='{eventId}'");

                var conversationIds = _currentEvent.ConversationIds
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                EventRecord? refreshedRecord;

                ProgressValue = 5;

                // Trigger catch-up for tracked conversations
                if (conversationIds.Count > 0)
                {
                    StatusMessage = string.Format(Properties.Resources.Syncing_conversationIds_Count_conversations, conversationIds.Count);
                    await _services.EventMonitor.TriggerCatchUpAsync(eventId, conversationIds, runImmediately: true, immediateTimeout: ManualCatchUpTimeout, useFullHistory: true).ConfigureAwait(false);
                }
                ProgressValue = 30;

                // Also trigger search for related subjects
                if (_currentEvent.RelatedSubjects != null && _currentEvent.RelatedSubjects.Count > 0)
                {
                    StatusMessage = Properties.Resources.Searching_for_related_subject_emails;
                    // This part is tricky because OutlookEventMonitor is designed around ConversationIDs.
                    // We might need to extend it or add a new method to search by Subject.
                    // For now, let's assume the user wants to rely on the "New Mail" event for subject matching,
                    // OR we can try to find mails in the current folder that match the subject.
                    // But since we are in "RefreshMailAsync", we should probably try to find missing mails.
                    
                    // However, implementing a full search here might be slow.
                    // Let's stick to the requested logic: "Subject集合中提到的标题需要在刷新时被寻找"
                    // We will add a method to OutlookEventMonitor to search by Subject.
                    // But since I cannot easily change OutlookEventMonitor's public interface without breaking things or reading it all,
                    // I will implement a local search here in the ViewModel or use a helper.
                    
                    // Actually, let's just rely on the fact that if we trigger a catch-up, it might find things if we pass the right IDs.
                    // But we don't have IDs for subject-matched mails yet.
                    
                    // Let's add a specific search task here.
                    await SearchBySubjectAsync(_currentEvent).ConfigureAwait(false);
                }
                ProgressValue = 50;

                StatusMessage = Properties.Resources.Reloading_event_data;
                refreshedRecord = await _services.EventRepository.GetByIdAsync(eventId).ConfigureAwait(false);
                ProgressValue = 60;

                if (refreshedRecord is null)
                {
                    DebugLogger.Log($"RefreshMailAsync unable to reload Event='{eventId}' after refresh");
                    return;
                }

                StatusMessage = Properties.Resources.Validating_email_availability;
                await ValidateMailEntriesAsync(refreshedRecord, 60, 90).ConfigureAwait(false);
                ProgressValue = 90;

                var finalRecord = await _services.EventRepository.GetByIdAsync(eventId).ConfigureAwait(false) ?? refreshedRecord;

                StatusMessage = Properties.Resources.Updating_UI;
                await _dispatcher.InvokeAsync(() =>
                {
                    // Guard against race condition: if user switched events while refresh was running, discard results
                    if (_currentEvent is null || !string.Equals(_currentEvent.EventId, eventId, StringComparison.OrdinalIgnoreCase))
                    {
                        DebugLogger.Log($"RefreshMailAsync discarded results for Event='{eventId}' because current event changed to '{_currentEvent?.EventId}'");
                        return;
                    }

                    _currentEvent = finalRecord;
                    PopulateMail(finalRecord);
                    PopulateAttachments(finalRecord);
                });
                ProgressValue = 100;

                DebugLogger.Log($"RefreshMailAsync completed for Event='{eventId}' MailCount={finalRecord.Emails.Count}");
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
                IsProgressIndeterminate = false;
                StatusMessage = Properties.Resources.Ready_1;
                ProgressValue = 0;
                RefreshMailCommand.RaiseCanExecuteChanged();
                RemoveMailCommand.RaiseCanExecuteChanged();
            }
        }

        private async Task<EventRecord?> WaitForCatchUpAsync(string eventId, IReadOnlyCollection<string> conversationIds, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(eventId))
            {
                return null;
            }

            var repository = _services.EventRepository;
            var monitor = _services.EventMonitor;
            if (monitor is null || conversationIds is null || conversationIds.Count == 0)
            {
                return await repository.GetByIdAsync(eventId).ConfigureAwait(false);
            }

            var normalized = conversationIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (normalized.Count == 0)
            {
                return await repository.GetByIdAsync(eventId).ConfigureAwait(false);
            }

            var tcs = new TaskCompletionSource<EventRecord?>(TaskCreationOptions.RunContinuationsAsynchronously);
            using var cts = new CancellationTokenSource(timeout);
            using var registration = cts.Token.Register(() => tcs.TrySetResult(null));
            string? observedReason = null;

            void Handler(object? sender, EventChangedEventArgs e)
            {
                if (!string.Equals(e.Record.EventId, eventId, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (!string.Equals(e.ChangeReason, "MailAppended", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(e.ChangeReason, "Updated", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (normalized.Count > 0 && !e.Record.ConversationIds.Any(id => normalized.Contains(id)))
                {
                    return;
                }

                observedReason = e.ChangeReason;
                tcs.TrySetResult(e.Record);
            }

            repository.EventChanged += Handler;

            try
            {
                DebugLogger.Log($"WaitForCatchUpAsync enqueued catch-up for Event='{eventId}' Conversations={normalized.Count}");
                await monitor.TriggerCatchUpAsync(eventId, normalized, runImmediately: true, immediateTimeout: timeout, useFullHistory: true).ConfigureAwait(false);

                var result = await tcs.Task.ConfigureAwait(false);
                if (result is not null)
                {
                    DebugLogger.Log($"WaitForCatchUpAsync observed change '{observedReason}' for Event='{eventId}'");
                    return result;
                }

                DebugLogger.Log($"WaitForCatchUpAsync timed out for Event='{eventId}' after {timeout.TotalSeconds:F1}s");
            }
            finally
            {
                repository.EventChanged -= Handler;
            }

            return await repository.GetByIdAsync(eventId).ConfigureAwait(false);
        }

        private IReadOnlyList<(string EntryId, string StoreId)> CollectConversationMailReferences(Application application, string conversationId, IReadOnlyCollection<string>? storeFilter)
        {
            var matches = new List<(string EntryId, string StoreId)>();
            if (application is null || string.IsNullOrWhiteSpace(conversationId))
            {
                return matches;
            }

            Stores? stores = null;
            ISet<string>? normalizedFilter = null;
            try
            {
                stores = application.Session?.Stores;
            }
            catch (COMException ex)
            {
                DebugLogger.Log($"CollectConversationMailReferences failed to access stores: {ex.Message}");
                stores = null;
            }
            catch
            {
                stores = null;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var cutoffUtc = DateTime.UtcNow.AddDays(-14);

            if (storeFilter is null || storeFilter.Count == 0)
            {
                Store? defaultStore = null;
                try
                {
                    defaultStore = application.Session?.DefaultStore;
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"CollectConversationMailReferences default store failed: {ex.Message}");
                    defaultStore = null;
                }

                if (defaultStore is null)
                {
                    return matches;
                }

                try
                {
                    var defaultStoreId = defaultStore.StoreID ?? string.Empty;
                    CollectConversationMatchesFromStore(defaultStore, defaultStoreId, conversationId, cutoffUtc, matches, seen);
                }
                finally
                {
                    ReleaseComObject(defaultStore);
                }

                return matches;
            }

            try
            {
                normalizedFilter = storeFilter
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Select(id => id.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"CollectConversationMailReferences filter normalization failed: {ex.Message}");
                normalizedFilter = null;
            }

            if (normalizedFilter is null || normalizedFilter.Count == 0 || stores is null)
            {
                return matches;
            }

            var remaining = new HashSet<string>(normalizedFilter, StringComparer.OrdinalIgnoreCase);

            try
            {
                for (var index = 1; index <= stores.Count; index++)
                {
                    Store? store = null;
                    try
                    {
                        store = stores[index];
                        if (store is null)
                        {
                            continue;
                        }

                        var storeId = store.StoreID ?? string.Empty;
                        if (!remaining.Contains(storeId))
                        {
                            continue;
                        }

                        remaining.Remove(storeId);
                        CollectConversationMatchesFromStore(store, storeId, conversationId, cutoffUtc, matches, seen);
                        if (remaining.Count == 0)
                        {
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(store);
                    }
                }
            }
            finally
            {
                ReleaseComObject(stores);
            }

            return matches;
        }

        private void CollectConversationMatchesFromStore(Store store, string storeId, string conversationId, DateTime cutoffUtc, IList<(string EntryId, string StoreId)> matches, ISet<string> seen)
        {
            if (store is null)
            {
                return;
            }

            MAPIFolder? inbox = null;
            MAPIFolder? sent = null;
            MAPIFolder? deleted = null;

            try
            {
                inbox = SafeGetDefaultFolder(store, OlDefaultFolders.olFolderInbox);
                if (inbox is not null)
                {
                    CollectMatchesFromFolder(inbox, storeId, conversationId, cutoffUtc, matches, seen, includeChildren: true);
                }

                sent = SafeGetDefaultFolder(store, OlDefaultFolders.olFolderSentMail);
                if (sent is not null)
                {
                    CollectMatchesFromFolder(sent, storeId, conversationId, cutoffUtc, matches, seen, includeChildren: false);
                }

                deleted = SafeGetDefaultFolder(store, OlDefaultFolders.olFolderDeletedItems);
                if (deleted is not null)
                {
                    CollectMatchesFromFolder(deleted, storeId, conversationId, cutoffUtc, matches, seen, includeChildren: false);
                }
            }
            finally
            {
                ReleaseComObject(deleted);
                ReleaseComObject(sent);
                ReleaseComObject(inbox);
            }
        }

        private void CollectMatchesFromFolder(MAPIFolder folder, string storeId, string conversationId, DateTime cutoffUtc, IList<(string EntryId, string StoreId)> matches, ISet<string> seen, bool includeChildren)
        {
            if (folder is null)
            {
                return;
            }

            try
            {
                if (folder.DefaultItemType != OlItemType.olMailItem)
                {
                    return;
                }
            }
            catch (COMException)
            {
                return;
            }
            catch
            {
                return;
            }

            Items? items = null;
            Items? filtered = null;
            object? current = null;

            try
            {
                items = folder.Items;
                if (items is null)
                {
                    return;
                }

                items.Sort("[ReceivedTime]", true);
                if (s_conversationFilterSupported)
                {
                    try
                    {
                        var conversationFilter = BuildConversationFilter(conversationId, cutoffUtc);
                        filtered = items.Restrict(conversationFilter);
                    }
                    catch (COMException ex)
                    {
                        s_conversationFilterSupported = false;
                        DebugLogger.Log($"CollectMatchesFromFolder conversation filter unsupported Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}. Falling back to received-time filter.");
                    }
                }

                if (filtered is null)
                {
                    var fallbackFilter = BuildReceivedSinceFilter(cutoffUtc);
                    filtered = items.Restrict(fallbackFilter);
                }

                if (filtered is null)
                {
                    return;
                }

                current = filtered.GetFirst();
                while (current is not null)
                {
                    if (current is MailItem mail)
                    {
                        try
                        {
                            var entryId = mail.EntryID;
                            var mailConversationId = string.Empty;
                            try
                            {
                                mailConversationId = mail.ConversationID ?? string.Empty;
                            }
                            catch (COMException)
                            {
                                mailConversationId = string.Empty;
                            }

                            if (!string.IsNullOrEmpty(entryId) &&
                                !string.IsNullOrEmpty(mailConversationId) &&
                                string.Equals(mailConversationId, conversationId, StringComparison.OrdinalIgnoreCase) &&
                                seen.Add(entryId))
                            {
                                matches.Add((entryId, storeId));
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                        }
                    }
                    else
                    {
                        Marshal.ReleaseComObject(current);
                    }

                    current = filtered.GetNext();
                }
            }
            catch (COMException ex)
            {
                DebugLogger.Log($"CollectMatchesFromFolder failed Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}");
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"CollectMatchesFromFolder unexpected exception Folder='{folder.Name}' Conversation='{conversationId}': {ex}");
            }
            finally
            {
                if (current is not null)
                {
                    Marshal.ReleaseComObject(current);
                }
                ReleaseComObject(filtered);
                ReleaseComObject(items);
            }

            if (!includeChildren)
            {
                return;
            }

            Folders? children = null;
            try
            {
                children = folder.Folders;
                if (children is null || children.Count == 0)
                {
                    return;
                }

                for (var index = 1; index <= children.Count; index++)
                {
                    MAPIFolder? child = null;
                    try
                    {
                        child = children[index];
                        if (child is null)
                        {
                            continue;
                        }

                        CollectMatchesFromFolder(child, storeId, conversationId, cutoffUtc, matches, seen, includeChildren: true);
                    }
                    finally
                    {
                        ReleaseComObject(child);
                    }
                }
            }
            finally
            {
                ReleaseComObject(children);
            }
        }

        private static string BuildConversationFilter(string conversationId, DateTime cutoffUtc)
        {
            var cutoffLocal = cutoffUtc.ToLocalTime();
            var escapedConversationId = EscapeForOutlookFilter(conversationId);
            return $"[ConversationID] = '{escapedConversationId}' AND [ReceivedTime] >= \"{cutoffLocal.ToString("g", OutlookFilterCulture)}\"";
        }

        private static string EscapeForOutlookFilter(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        private static string BuildReceivedSinceFilter(DateTime cutoffUtc)
        {
            var cutoffLocal = cutoffUtc.ToLocalTime();
            return $"[ReceivedTime] >= \"{cutoffLocal.ToString("g", OutlookFilterCulture)}\"";
        }

        private async Task ValidateMailEntriesAsync(EventRecord record, double startProgress = -1, double endProgress = -1)
        {
            if (record is null)
            {
                return;
            }

            if (record.Emails.Count == 0)
            {
                return;
            }

            var application = _services.OutlookApplication;
            if (application is null)
            {
                return;
            }

            var emails = record.Emails.Where(e => !e.IsRemoved).ToList();
            int total = emails.Count;
            int current = 0;
            bool reportProgress = startProgress >= 0 && endProgress > startProgress;

            foreach (var email in emails)
            {
                current++;
                if (reportProgress)
                {
                    double p = startProgress + (endProgress - startProgress) * ((double)current / total);
                    ProgressValue = p;
                }

                MailItem? resolved = null;
                var wasNew = email.IsNewOrUpdated;

                try
                {
                    resolved = ResolveMail(application, email.EntryId, email.StoreId, email.InternetMessageId);
                    if (resolved is null)
                    {
                        DebugLogger.Log($"Mail validation failed for Event='{record.EventId}' Entry='{email.EntryId}' MsgId='{email.InternetMessageId}'");
                        continue;
                    }

                    var resolvedEntryId = resolved.EntryID ?? string.Empty;
                    var resolvedStoreId = GetStoreId(resolved);
                    var entryMismatch = !string.Equals(resolvedEntryId, email.EntryId, StringComparison.OrdinalIgnoreCase);
                    var storeMismatch = !string.IsNullOrEmpty(resolvedStoreId) && !string.Equals(resolvedStoreId, email.StoreId, StringComparison.OrdinalIgnoreCase);

                    if (entryMismatch || storeMismatch)
                    {
                        DebugLogger.Log($"Mail validation refreshing metadata for Event='{record.EventId}' MsgId='{email.InternetMessageId}' NewEntry='{resolvedEntryId}'");
                        var updatedRecord = await _services.EventRepository.AddMailToEventAsync(record.EventId, resolved).ConfigureAwait(false);
                        if (updatedRecord is not null)
                        {
                            var refreshedMail = updatedRecord.Emails.FirstOrDefault(e =>
                                (!string.IsNullOrEmpty(email.InternetMessageId) && string.Equals(e.InternetMessageId, email.InternetMessageId, StringComparison.OrdinalIgnoreCase)) ||
                                (!string.IsNullOrEmpty(resolvedEntryId) && string.Equals(e.EntryId, resolvedEntryId, StringComparison.OrdinalIgnoreCase)));

                            if (refreshedMail is not null)
                            {
                                refreshedMail.IsNewOrUpdated = wasNew;
                                await _services.EventRepository.UpdateAsync(updatedRecord).ConfigureAwait(false);
                            }
                        }
                    }
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"Mail validation COMException for Event='{record.EventId}' Entry='{email.EntryId}': {ex.Message}");
                }
                finally
                {
                    ReleaseComObject(resolved);
                }
            }
        }

        private static string GetStoreId(MailItem mailItem)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            MAPIFolder? parent = null;
            try
            {
                parent = mailItem.Parent as MAPIFolder;
                return parent?.StoreID ?? string.Empty;
            }
            catch (COMException)
            {
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(parent);
            }
        }

        private static void ReleaseComObject(object? comObject)
        {
            if (comObject is null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // ignore release errors
            }
        }

        private void InvokeOnUiThread(System.Action action)
        {
            if (action is null)
            {
                return;
            }

            if (_dispatcher.CheckAccess())
            {
                action();
            }
            else
            {
                _dispatcher.Invoke(action);
            }
        }

        private async Task SearchBySubjectAsync(EventRecord record)
        {
            if (record is null) return;

            var application = _services.OutlookApplication;
            if (application is null) return;

            await Task.Run(async () =>
            {
                try
                {
                    var session = application.Session;
                    var folders = new List<MAPIFolder>();
                    
                    try { folders.Add(session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)); } catch {}
                    try { folders.Add(session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail)); } catch {}
                    try { folders.Add(session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)); } catch {}

                    foreach (var folder in folders)
                    {
                        if (folder != null)
                        {
                            await SearchFolderForEventAsync(folder, record);
                            Marshal.ReleaseComObject(folder);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLogger.Log($"SearchBySubjectAsync failed: {ex.Message}");
                }
            });
        }

        private async Task SearchFolderForEventAsync(MAPIFolder folder, EventRecord record)
        {
             Items? items = null;
             Items? restricted = null;
             try 
             {
                 items = folder.Items;
                 // Filter by time to reduce load (Broad casting)
                 var filter = $"[ReceivedTime] >= \"{DateTime.Now.AddDays(-SubjectSearchRangeDays).ToString("g", OutlookFilterCulture)}\"";
                 restricted = items.Restrict(filter);
                 
                 // We iterate manually to release objects
                 object? item = restricted.GetFirst();
                 while (item != null)
                 {
                     if (item is MailItem mail)
                     {
                         try
                         {
                             // Optimization: Check if already in event by EntryID (fastest)
                             if (!record.Emails.Any(e => e.EntryId == mail.EntryID))
                             {
                                 bool isMatch = false;

                                 // 1. Subject Match (Fastest)
                                 var normSubject = NormalizeSubject(mail.Subject);
                                 if (record.RelatedSubjects != null && record.RelatedSubjects.Contains(normSubject))
                                 {
                                     // Check participants to be safe
                                     var participants = MailParticipantExtractor.Capture(mail);
                                     if (MailParticipantExtractor.Intersects(participants, record.Participants))
                                     {
                                         isMatch = true;
                                     }
                                 }

                                 // 2. In-Reply-To Match (Fast & Accurate)
                                 if (!isMatch)
                                 {
                                     try 
                                     {
                                         var pa = mail.PropertyAccessor;
                                         // PR_IN_REPLY_TO_ID = 0x1042001F
                                         string? inReplyTo = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1042001F") as string;
                                         if (!string.IsNullOrEmpty(inReplyTo) && record.ProcessedMessageIds != null && record.ProcessedMessageIds.Contains(inReplyTo!))
                                         {
                                             isMatch = true;
                                         }
                                     }
                                     catch {} // Property might not exist or error accessing it
                                 }

                                 if (isMatch)
                                 {
                                     await _services.EventRepository.AddMailToEventAsync(record.EventId, mail).ConfigureAwait(false);
                                 }
                             }
                         }
                         finally
                         {
                             Marshal.ReleaseComObject(mail);
                         }
                     }
                     else
                     {
                         Marshal.ReleaseComObject(item);
                     }
                     
                     item = restricted.GetNext();
                 }
             }
             catch (System.Exception ex)
             {
                 DebugLogger.Log($"SearchFolderForEventAsync failed for folder {folder.Name}: {ex.Message}");
             }
             finally
             {
                 if (restricted != null) Marshal.ReleaseComObject(restricted);
                 if (items != null) Marshal.ReleaseComObject(items);
             }
        }

        private static string NormalizeSubject(string? subject)
        {
            if (string.IsNullOrWhiteSpace(subject)) return string.Empty;
            var normalized = subject!.Trim();
            var prefixes = new[] { "RE:", "FW:", "FWD:", "回复:", "转发:", "回覆:", "轉寄:" };
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (var prefix in prefixes)
                {
                    if (normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        normalized = normalized.Substring(prefix.Length).Trim();
                        changed = true;
                    }
                }
            }
            return normalized;
        }

        private async Task MarkAllAsReadAsync()
        {
            if (_currentEvent is null) return;

            var changed = false;
            foreach (var email in _currentEvent.Emails)
            {
                if (email.IsNewOrUpdated)
                {
                    email.IsNewOrUpdated = false;
                    changed = true;

                    // Fix: Add to ProcessedMessageIds to prevent re-highlighting on refresh
                    if (!string.IsNullOrEmpty(email.InternetMessageId))
                    {
                        if (_currentEvent.ProcessedMessageIds == null)
                        {
                            _currentEvent.ProcessedMessageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        }
                        _currentEvent.ProcessedMessageIds.Add(email.InternetMessageId);
                    }
                }
            }

            if (changed)
            {
                await _services.EventRepository.UpdateAsync(_currentEvent);
                // Refresh UI to reflect changes (remove highlights)
                PopulateMail(_currentEvent);
            }
        }

        private void OpenTemplateFile(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
            {
                return;
            }
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch { }
        }

        private void OnExplorerSelectionChange()
        {
            // If selection changes and we are NOT currently setting it programmatically,
            // it means the user (or another process) changed the selection.
            // We should cancel any pending preview to avoid "rolling back" the user's selection.
            if (!_isSettingSelection)
            {
                CancelPreview();

                try
                {
                    if (_monitoredExplorer != null)
                    {
                        var selection = _monitoredExplorer.Selection;
                        try
                        {
                            if (selection.Count == 1)
                            {
                                var item = selection[1];
                                if (item is MailItem mailItem)
                                {
                                    try
                                    {
                                        var entryId = mailItem.EntryID;
                                        var match = MailItems.FirstOrDefault(m => m.EntryId == entryId);

                                        if (match != null && match != SelectedMail)
                                        {
                                            _suppressPreview = true;
                                            try
                                            {
                                                SelectedMail = match;
                                            }
                                            finally
                                            {
                                                _suppressPreview = false;
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(mailItem);
                                    }
                                }
                                else if (item != null)
                                {
                                    Marshal.ReleaseComObject(item);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(selection);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLogger.Log($"OnExplorerSelectionChange sync failed: {ex.Message}");
                }
            }
        }

        public void CancelPreview()
        {
            _previewCts?.Cancel();
            _previewCts = null;
        }

        private async Task RemoveTemplateFileAsync(object? param)
        {
            if (param is not System.Collections.IList items || items.Count == 0)
            {
                return;
            }

            var filesToRemove = items.Cast<TemplateFileViewModel>().ToList();
            bool eventChanged = false;
            bool templateChanged = false;

            IsBusy = true;
            _services.BusyState.SetBusy("Removing files...");
            try
            {
                foreach (var file in filesToRemove)
                {
                    // 1. Try removing from Event specific files
                    // Use OriginalPath if available, else FilePath
                    string pathToCheck = !string.IsNullOrEmpty(file.OriginalPath) ? file.OriginalPath : file.FilePath;

                    if (_currentEvent?.AdditionalFiles != null && _currentEvent.AdditionalFiles.Contains(pathToCheck))
                    {
                        _currentEvent.AdditionalFiles.Remove(pathToCheck);
                        eventChanged = true;
                    }
                    
                    // 2. Try removing from Template files (by excluding them)
                    if (SelectedTemplate?.AttachmentPaths != null && SelectedTemplate.AttachmentPaths.Contains(pathToCheck))
                    {
                        if (_currentEvent != null)
                        {
                            if (_currentEvent.ExcludedTemplateFiles == null)
                            {
                                _currentEvent.ExcludedTemplateFiles = new List<string>();
                            }
                            if (!_currentEvent.ExcludedTemplateFiles.Contains(pathToCheck))
                            {
                                _currentEvent.ExcludedTemplateFiles.Add(pathToCheck);
                                eventChanged = true;
                            }
                        }
                    }
                }

                if (eventChanged && _currentEvent != null)
                {
                    await _services.EventRepository.UpdateAsync(_currentEvent);
                }

                if (templateChanged && SelectedTemplate != null)
                {
                    _services.DashboardTemplates.AddOrUpdateTemplate(SelectedTemplate);
                }
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
            }

            if (eventChanged || templateChanged)
            {
                UpdateTemplateFiles();
            }
        }

        private async Task GenerateFolderAsync()
        {
            if (_currentEvent is null) return;

            var eventTitleSafe = string.Join("_", _currentEvent.EventTitle.Split(Path.GetInvalidFileNameChars()));

            using (var dialog = new WinForms.SaveFileDialog())
            {
                dialog.Title = Properties.Resources.Create_Event_Folder_Enter_folder_name;
                dialog.FileName = eventTitleSafe;
                dialog.Filter = Properties.Resources.Folder;
                dialog.CheckFileExists = false;
                dialog.OverwritePrompt = false;
                dialog.AddExtension = false;
                
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    var folderPath = dialog.FileName;
                    
                    try
                    {
                        IsBusy = true;
                        _services.BusyState.SetBusy("Creating folder...");
                        if (!Directory.Exists(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        WinForms.MessageBox.Show(string.Format(Properties.Resources.Failed_to_create_folder_ex_Message, ex.Message), Properties.Resources.Error, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                        return;
                    }
                    finally
                    {
                        IsBusy = false;
                        _services.BusyState.ClearBusy();
                    }

                    LocalFolderPath = folderPath;
                    if (_currentEvent != null)
                    {
                         await _services.EventRepository.UpdateAsync(_currentEvent);
                    }
                    
                    await UpdateFolderAsync();
                               }
            }
        }

        private async Task UpdateFolderAsync()
        {
            if (_currentEvent is null || string.IsNullOrEmpty(LocalFolderPath)) return;

            if (!Directory.Exists(LocalFolderPath))
            {
                WinForms.MessageBox.Show(string.Format(Properties.Resources.Folder_does_not_exist_LocalFolderPath, LocalFolderPath), Properties.Resources.Error, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                return;
            }

            IsBusy = true;
            _services.BusyState.SetBusy(Properties.Resources.Updating_folder);
            StatusMessage = Properties.Resources.Updating_folder;
            ProgressValue = 0;
            IsProgressIndeterminate = false;

            bool eventChanged = false;

            try
            {
                var filesToProcess = TemplateFiles.ToList();
                var total = filesToProcess.Count;
                int current = 0;

                await Task.Run(() => 
                {
                    foreach (var fileVm in filesToProcess)
                    {
                        current++;
                        _dispatcher.Invoke(() => 
                        {
                            ProgressValue = (double)current / total * 100;
                            StatusMessage = string.Format(Properties.Resources.Processing_current_total_fileVm_FileName, current, total, fileVm.FileName);
                        });

                        var sourcePath = fileVm.FilePath;
                        if (!File.Exists(sourcePath)) continue;

                        // Skip if already in the target folder
                        if (Path.GetDirectoryName(sourcePath)?.Equals(LocalFolderPath, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            continue;
                        }

                        var fileName = Path.GetFileName(sourcePath);
                        var targetPath = Path.Combine(LocalFolderPath, fileName);

                        try 
                        {
                            if (!File.Exists(targetPath))
                            {
                                File.Copy(sourcePath, targetPath);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            DebugLogger.Log($"Failed to copy file {fileName}: {ex.Message}");
                        }
                    }
                });

                // Update paths in AdditionalFiles to point to the new folder
                if (_currentEvent.AdditionalFiles != null)
                {
                    var newFiles = new List<string>();
                    foreach (var path in _currentEvent.AdditionalFiles)
                    {
                        var fileName = Path.GetFileName(path);
                        var targetPath = Path.Combine(LocalFolderPath, fileName);
                        
                        // If the file exists in the target folder (we just copied it or it was there), update the path
                        if (File.Exists(targetPath))
                        {
                            newFiles.Add(targetPath);
                                                       if (!string.Equals(path, targetPath, StringComparison.OrdinalIgnoreCase))
                            {
                                eventChanged = true;
                            }
                        }
                        else
                        {
                            newFiles.Add(path);
                        }
                    }
                    _currentEvent.AdditionalFiles = newFiles;
                }



                if (eventChanged)
                {
                    await _services.EventRepository.UpdateAsync(_currentEvent);
                    _dispatcher.Invoke(UpdateTemplateFiles);
                }
                
                StatusMessage = Properties.Resources.Folder_update_completed;
            }
            catch (System.Exception ex)
            {
                WinForms.MessageBox.Show(string.Format(Properties.Resources.Failed_to_update_folder_ex_Message, ex.Message), Properties.Resources.Error, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                StatusMessage = Properties.Resources.Update_Failed;
            }
            finally
            {
                IsBusy = false;
                _services.BusyState.ClearBusy();
                await Task.Delay(2000);
                StatusMessage = Properties.Resources.Ready_1;
            }
        }

        private void UpdateTemplateFiles()
        {
            var list = new ObservableCollection<TemplateFileViewModel>();
            
            // 1. Add files from the Template
            if (SelectedTemplate != null)
            {
                if (SelectedTemplate.AttachmentPaths == null)
                {
                    SelectedTemplate.AttachmentPaths = new List<string>();
                }

                foreach (var path in SelectedTemplate.AttachmentPaths)
                {
                    // Check if excluded
                    if (_currentEvent?.ExcludedTemplateFiles != null && _currentEvent.ExcludedTemplateFiles.Contains(path))
                    {
                        continue;
                    }

                    string localPath = EnsureFileLocal(path);
                    var vm = new TemplateFileViewModel(localPath);
                    vm.OriginalPath = path;
                    vm.IsCommonFile = true;
                    list.Add(vm);
                }
            }

            // 2. Add files specific to this Event
            if (_currentEvent != null)
            {
                if (_currentEvent.AdditionalFiles == null)
                {
                    _currentEvent.AdditionalFiles = new List<string>();
                }

                foreach (var path in _currentEvent.AdditionalFiles)
                {
                    string localPath = EnsureFileLocal(path);
                    
                    // Avoid duplicates if the file is already in the template
                    if (!list.Any(x => string.Equals(x.OriginalPath, path, StringComparison.OrdinalIgnoreCase) || string.Equals(x.FilePath, localPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        var vm = new TemplateFileViewModel(localPath);
                        vm.OriginalPath = path;
                        list.Add(vm);
                    }
                }
            }

            TemplateFiles = list;
            DebugLogger.Log($"Updated TemplateFiles. Count: {TemplateFiles.Count}");
        }

        public void AddTemplateFiles(IEnumerable<string> files)
        {
            if (_currentEvent == null) return;
            
            if (_currentEvent.AdditionalFiles == null)
            {
                _currentEvent.AdditionalFiles = new List<string>();
            }

            bool changed = false;
            foreach (var file in files)
            {
                if (!_currentEvent.AdditionalFiles.Contains(file, StringComparer.OrdinalIgnoreCase))
                {
                    _currentEvent.AdditionalFiles.Add(file);
                    changed = true;
                }
            }
            
            if (changed)
            {
                _ = _services.EventRepository.UpdateAsync(_currentEvent);
                UpdateTemplateFiles();
            }
        }

        private void CopyJsonKeys()
        {
            try
            {
                var items = DashboardItems
                    .Where(item => !string.IsNullOrWhiteSpace(item.Value))
                    .Select(item => $"{item.Key}: {item.Value}");

                var text = string.Join(Environment.NewLine, items);
                if (!string.IsNullOrEmpty(text))
                {
                    System.Windows.Clipboard.SetText(text);
                    StatusMessage = Properties.Resources.Copied_non_empty_Json_keys_to_clipboard;
                }
                else
                {
                    StatusMessage = Properties.Resources.No_non_empty_Json_keys_to_copy;
                }
            }
            catch (System.Exception ex)
            {
                StatusMessage = string.Format(Properties.Resources.Copy_failed_ex_Message, ex.Message);
            }
        }

        private string EnsureFileLocal(string originalPath)
        {
            if (_currentEvent == null) return originalPath;
            
            string storageRoot = Properties.Settings.Default.EventFilesStoragePath;
            if (string.IsNullOrWhiteSpace(storageRoot))
            {
                storageRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OSEMAddIn");
            }
            
            string eventFolder = Path.Combine(storageRoot, "documents", _currentEvent.EventId);
            try
            {
                if (!Directory.Exists(eventFolder))
                {
                    Directory.CreateDirectory(eventFolder);
                }
                
                string fileName = Path.GetFileName(originalPath);
                string localPath = Path.Combine(eventFolder, fileName);
                
                if (!File.Exists(localPath) && File.Exists(originalPath))
                {
                    File.Copy(originalPath, localPath, true);
                }
                
                return File.Exists(localPath) ? localPath : originalPath;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"Failed to ensure local file for {originalPath}: {ex.Message}");
                return originalPath;
            }
        }
    }
}
