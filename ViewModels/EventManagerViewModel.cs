#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Office.Interop.Outlook;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;
using OSEMAddIn.Services;
using OSEMAddIn.Views.EventManager;

namespace OSEMAddIn.ViewModels
{
    internal sealed class EventManagerViewModel : ViewModelBase
    {
        private readonly ServiceContainer _services;
    private readonly ObservableCollection<EventListItemViewModel> _ongoingItems = new();
    private readonly ObservableCollection<EventListItemViewModel> _archivedItems = new();
    private readonly Dispatcher _dispatcher;
        private string _searchTerm = string.Empty;
        private DashboardTemplate? _selectedFilterTemplate;
        private bool _isBusy;
        private TemplateEditorView? _openedEditor;
        private System.Threading.CancellationTokenSource? _refreshCts;
        private string _refreshStatusText = string.Empty;
        private double _refreshProgressValue;
        private bool _isRefreshing;
        private DateTime? _filterStartDate;
        private DateTime? _filterEndDate;
        private int _selectedTabIndex;

        public EventManagerViewModel(ServiceContainer services)
        {
            _services = services ?? throw new ArgumentNullException(nameof(services));
            _dispatcher = Dispatcher.CurrentDispatcher;
            _services.EventRepository.EventChanged += OnEventChanged;

            var now = DateTime.Now;
            _filterStartDate = new DateTime(now.Year, now.Month, 1);
            _filterEndDate = _filterStartDate.Value.AddMonths(1).AddDays(-1);

            OngoingEvents = CollectionViewSource.GetDefaultView(_ongoingItems);
            OngoingEvents.Filter = FilterOngoing;
            ((INotifyCollectionChanged)OngoingEvents.SortDescriptions).CollectionChanged += (s, e) =>
            {
                if (OngoingEvents.SortDescriptions.Count > 0)
                {
                    SavedOngoingSort = OngoingEvents.SortDescriptions[0];
                }
            };

            ArchivedEvents = CollectionViewSource.GetDefaultView(_archivedItems);
            ArchivedEvents.Filter = FilterArchived;
            ((INotifyCollectionChanged)ArchivedEvents.SortDescriptions).CollectionChanged += (s, e) =>
            {
                if (ArchivedEvents.SortDescriptions.Count > 0)
                {
                    SavedArchivedSort = ArchivedEvents.SortDescriptions[0];
                }
            };

            RefreshCommand = new AsyncRelayCommand(_ => ReloadAsync());
            BatchArchiveCommand = new AsyncRelayCommand(_ => ArchiveSelectionAsync(), _ => SelectedEventIds.Count > 0);
            DeleteCommand = new AsyncRelayCommand(_ => DeleteSelectionAsync(), _ => SelectedEventIds.Count > 0);
            ArchiveSingleCommand = new AsyncRelayCommand(param => ArchiveSingleAsync(param as EventListItemViewModel));
            ReopenEventCommand = new AsyncRelayCommand(param => ReopenEventAsync(param as EventListItemViewModel));
            OpenEventCommand = new RelayCommand(param => OpenEvent(param as EventListItemViewModel));
            RebuildExportFilterCommand = new RelayCommand(_ => ArchivedEvents.Refresh());
            HandleMailDropCommand = new AsyncRelayCommand(_ => HandleMailDropAsync());
            RefreshAllEventsCommand = new AsyncRelayCommand(_ => RefreshAllEventsAsync());
            OpenRuleManagerCommand = new RelayCommand(_ => OpenRuleManager());
            OpenTemplateEditorCommand = new RelayCommand(_ => OpenTemplateEditor());
            SetFilterToNullCommand = new RelayCommand(_ => SelectedFilterTemplate = null);
            ChangeEventTitleCommand = new AsyncRelayCommand(param => ChangeEventTitleAsync(param as EventListItemViewModel));
            ClearSearchCommand = new RelayCommand(_ => SearchTerm = string.Empty);
            ResetDateFilterCommand = new RelayCommand(_ => ResetDateFilter());
            
            ExecuteGlobalScriptCommand = new AsyncRelayCommand(_ => ExecuteGlobalScriptAsync(), _ => SelectedGlobalScript != null);
            LoadGlobalScripts();
        }

        public event EventHandler<string>? OpenEventRequested;

        public ICollectionView OngoingEvents { get; }
        public ICollectionView ArchivedEvents { get; }
        public AsyncRelayCommand RefreshCommand { get; }
        public AsyncRelayCommand BatchArchiveCommand { get; }
        public AsyncRelayCommand DeleteCommand { get; }
        public AsyncRelayCommand ArchiveSingleCommand { get; }
        public AsyncRelayCommand ReopenEventCommand { get; }
        public RelayCommand OpenEventCommand { get; }
        public RelayCommand RebuildExportFilterCommand { get; }
        public AsyncRelayCommand HandleMailDropCommand { get; }
        public AsyncRelayCommand RefreshAllEventsCommand { get; }
        public RelayCommand OpenRuleManagerCommand { get; }
        public RelayCommand OpenTemplateEditorCommand { get; }
        public RelayCommand SetFilterToNullCommand { get; }
        public AsyncRelayCommand ChangeEventTitleCommand { get; }
        public RelayCommand ClearSearchCommand { get; }
        public RelayCommand ResetDateFilterCommand { get; }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public SortDescription? SavedOngoingSort { get; set; }
        public SortDescription? SavedArchivedSort { get; set; }

        private double _ongoingScrollOffset;
        public double OngoingScrollOffset
        {
            get => _ongoingScrollOffset;
            set => SetProperty(ref _ongoingScrollOffset, value);
        }

        private double _archivedScrollOffset;
        public double ArchivedScrollOffset
        {
            get => _archivedScrollOffset;
            set => SetProperty(ref _archivedScrollOffset, value);
        }

        public DateTime? FilterStartDate
        {
            get => _filterStartDate;
            set
            {
                if (SetProperty(ref _filterStartDate, value))
                {
                    ArchivedEvents?.Refresh();
                }
            }
        }

        public DateTime? FilterEndDate
        {
            get => _filterEndDate;
            set
            {
                if (SetProperty(ref _filterEndDate, value))
                {
                    ArchivedEvents?.Refresh();
                }
            }
        }

        private void ResetDateFilter()
        {
            var now = DateTime.Now;
            FilterStartDate = new DateTime(now.Year, now.Month, 1);
            FilterEndDate = FilterStartDate.Value.AddMonths(1).AddDays(-1);
        }
        
        public AsyncRelayCommand ExecuteGlobalScriptCommand { get; }
        public ObservableCollection<PythonScriptDefinition> GlobalScripts { get; } = new();
        
        private PythonScriptDefinition? _selectedGlobalScript;
        public PythonScriptDefinition? SelectedGlobalScript
        {
            get => _selectedGlobalScript;
            set
            {
                if (SetProperty(ref _selectedGlobalScript, value))
                {
                    ExecuteGlobalScriptCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public ObservableCollection<string> SelectedEventIds { get; } = new();

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
                RaisePropertyChanged();
            }
        }

        public string SearchTerm
        {
            get => _searchTerm;
            set
            {
                if (_searchTerm == value)
                {
                    return;
                }

                _searchTerm = value ?? string.Empty;
                RaisePropertyChanged();
                OngoingEvents.Refresh();
                ArchivedEvents.Refresh();
            }
        }

        public DashboardTemplate? SelectedFilterTemplate
        {
            get => _selectedFilterTemplate;
            set
            {
                if (_selectedFilterTemplate == value)
                {
                    return;
                }

                _selectedFilterTemplate = value;
                RaisePropertyChanged();
                OngoingEvents.Refresh();
                ArchivedEvents.Refresh();
            }
        }

        public IReadOnlyList<DashboardTemplate> DashboardTemplates => _services.DashboardTemplates.GetTemplates();

        public async Task<string?> ExportArchivedAsync(string targetDirectory, string fileName)
        {
            if (SelectedFilterTemplate is null)
            {
                return null;
            }

            var records = await _services.EventRepository.GetAllAsync();
            var filtered = records.Where(r => r.Status == EventStatus.Archived && string.Equals(r.DashboardTemplateId, SelectedFilterTemplate.TemplateId, StringComparison.OrdinalIgnoreCase)).ToList();
            if (filtered.Count == 0)
            {
                return null;
            }

            return await _services.CsvExport.ExportAsync(filtered, SelectedFilterTemplate, targetDirectory, fileName);
        }

        public async Task InitializeAsync()
        {
            await ReloadAsync();
        }

        public void UpdateSelection(IList<object> selectedItems)
        {
            SelectedEventIds.Clear();
            foreach (var item in selectedItems.OfType<EventListItemViewModel>())
            {
                if (!SelectedEventIds.Contains(item.EventId))
                {
                    SelectedEventIds.Add(item.EventId);
                }
            }

            BatchArchiveCommand.RaiseCanExecuteChanged();
            DeleteCommand.RaiseCanExecuteChanged();
        }

        private async Task ReloadAsync()
        {
            if (IsBusy)
            {
                return;
            }

            IsBusy = true;
            try
            {
                var records = await _services.EventRepository.GetAllAsync();
                UpdateList(_ongoingItems, records.Where(r => r.Status == EventStatus.Open));
                UpdateList(_archivedItems, records.Where(r => r.Status == EventStatus.Archived));
                OngoingEvents.Refresh();
                ArchivedEvents.Refresh();
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void UpdateList(ObservableCollection<EventListItemViewModel> target, IEnumerable<EventRecord> records)
        {
            var existingLookup = target.ToDictionary(item => item.EventId, StringComparer.OrdinalIgnoreCase);
            var incomingIds = new HashSet<string>(records.Select(r => r.EventId), StringComparer.OrdinalIgnoreCase);

            foreach (var record in records)
            {
                if (existingLookup.TryGetValue(record.EventId, out var viewModel))
                {
                    viewModel.Update(record);
                }
                else
                {
                    target.Add(new EventListItemViewModel(record, _services.EventRepository));
                }
            }

            for (var i = target.Count - 1; i >= 0; i--)
            {
                var item = target[i];
                if (!incomingIds.Contains(item.EventId))
                {
                    target.RemoveAt(i);
                }
            }
        }

        private async Task ArchiveSelectionAsync()
        {
            if (SelectedEventIds.Count == 0)
            {
                return;
            }

            await _services.EventRepository.ArchiveAsync(SelectedEventIds);
            SelectedEventIds.Clear();
            BatchArchiveCommand.RaiseCanExecuteChanged();
        }

        private async Task ArchiveSingleAsync(EventListItemViewModel? item)
        {
            if (item is null)
            {
                return;
            }

            await _services.EventRepository.ArchiveAsync(new[] { item.EventId });
        }

        private async Task ReopenEventAsync(EventListItemViewModel? item)
        {
            if (item is null)
            {
                DebugLogger.Log("ReopenEventAsync called with null item.");
                return;
            }

            DebugLogger.Log($"ReopenEventAsync called for {item.EventId}");
            try
            {
                await _services.EventRepository.ReopenAsync(item.EventId);
                DebugLogger.Log($"ReopenAsync completed for {item.EventId}");
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"ReopenAsync failed: {ex}");
                System.Windows.MessageBox.Show($"Reopen failed: {ex.Message}");
            }
        }

        private void OpenEvent(EventListItemViewModel? item)
        {
            if (item is null)
            {
                return;
            }

            item.MarkMailAsRead();
            OpenEventRequested?.Invoke(this, item.EventId);
        }

        private bool FilterOngoing(object obj)
        {
            if (obj is not EventListItemViewModel item)
            {
                return false;
            }

            if (SelectedFilterTemplate != null)
            {
                bool match = string.Equals(item.DashboardTemplateId, SelectedFilterTemplate.TemplateId, StringComparison.OrdinalIgnoreCase);
                if (!match && string.IsNullOrEmpty(item.DashboardTemplateId) && string.Equals(SelectedFilterTemplate.TemplateId, "GEN", StringComparison.OrdinalIgnoreCase))
                {
                    match = true;
                }
                if (!match) return false;
            }

            if (string.IsNullOrWhiteSpace(SearchTerm))
            {
                return true;
            }

            return item.EventTitle.IndexOf(SearchTerm, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   item.EventId.IndexOf(SearchTerm, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool FilterArchived(object obj)
        {
            if (obj is not EventListItemViewModel item)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(SearchTerm))
            {
                bool matchesSearch = item.EventTitle.IndexOf(SearchTerm, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                     item.EventId.IndexOf(SearchTerm, StringComparison.OrdinalIgnoreCase) >= 0;
                if (!matchesSearch)
                {
                    return false;
                }
            }

            if (FilterStartDate.HasValue && item.LastUpdatedOn.Date < FilterStartDate.Value.Date)
            {
                return false;
            }
            if (FilterEndDate.HasValue && item.LastUpdatedOn.Date > FilterEndDate.Value.Date)
            {
                return false;
            }

            if (SelectedFilterTemplate is null)
            {
                return true;
            }

            bool match = string.Equals(item.DashboardTemplateId, SelectedFilterTemplate.TemplateId, StringComparison.OrdinalIgnoreCase);
            if (!match && string.IsNullOrEmpty(item.DashboardTemplateId) && string.Equals(SelectedFilterTemplate.TemplateId, "GEN", StringComparison.OrdinalIgnoreCase))
            {
                match = true;
            }
            return match;
        }

        private async Task ChangeEventTitleAsync(EventListItemViewModel? item)
        {
            if (item == null) return;

            var inputWindow = new System.Windows.Window
            {
                Title = "Change Event Title",
                Width = 400,
                Height = 150,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                ResizeMode = System.Windows.ResizeMode.NoResize
            };

            var stackPanel = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(10) };
            var textBox = new System.Windows.Controls.TextBox { Text = item.EventTitle, Margin = new System.Windows.Thickness(0, 0, 0, 10) };
            var buttonPanel = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Right };
            var okButton = new System.Windows.Controls.Button { Content = "OK", Width = 75, IsDefault = true };
            var cancelButton = new System.Windows.Controls.Button { Content = "Cancel", Width = 75, Margin = new System.Windows.Thickness(10, 0, 0, 0), IsCancel = true };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            stackPanel.Children.Add(new System.Windows.Controls.TextBlock { Text = "New Title:", Margin = new System.Windows.Thickness(0, 0, 0, 5) });
            stackPanel.Children.Add(textBox);
            stackPanel.Children.Add(buttonPanel);
            inputWindow.Content = stackPanel;

            bool? result = false;
            okButton.Click += (s, e) => { result = true; inputWindow.Close(); };
            cancelButton.Click += (s, e) => { inputWindow.Close(); };

            inputWindow.ShowDialog();

            if (result == true && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text != item.EventTitle)
            {
                var newTitle = textBox.Text.Trim();
                var eventRecord = await _services.EventRepository.GetByIdAsync(item.EventId);
                if (eventRecord != null)
                {
                    eventRecord.EventTitle = newTitle;
                    eventRecord.LastUpdatedOn = DateTime.UtcNow;
                    await _services.EventRepository.UpdateAsync(eventRecord);
                }
            }
        }

        private void OnEventChanged(object? sender, EventChangedEventArgs e)
        {
            DebugLogger.Log($"OnEventChanged: {e.Record.EventId}, Reason: {e.ChangeReason}, Status: {e.Record.Status}");

            if (!_dispatcher.CheckAccess())
            {
                // Marshal back to the UI thread before mutating bound collections.
                _dispatcher.BeginInvoke(new System.Action(() => OnEventChanged(sender, e)));
                return;
            }

            if (string.Equals(e.ChangeReason, "Deleted", StringComparison.OrdinalIgnoreCase))
            {
                var ongoing = _ongoingItems.FirstOrDefault(vm => string.Equals(vm.EventId, e.Record.EventId, StringComparison.OrdinalIgnoreCase));
                if (ongoing != null) _ongoingItems.Remove(ongoing);

                var archived = _archivedItems.FirstOrDefault(vm => string.Equals(vm.EventId, e.Record.EventId, StringComparison.OrdinalIgnoreCase));
                if (archived != null) _archivedItems.Remove(archived);

                OngoingEvents.Refresh();
                ArchivedEvents.Refresh();
                return;
            }

            var targetList = e.Record.Status == EventStatus.Open ? _ongoingItems : _archivedItems;
            var otherList = e.Record.Status == EventStatus.Open ? _archivedItems : _ongoingItems;

            var viewModel = targetList.FirstOrDefault(vm => string.Equals(vm.EventId, e.Record.EventId, StringComparison.OrdinalIgnoreCase));
            if (viewModel is null)
            {
                targetList.Add(new EventListItemViewModel(e.Record, _services.EventRepository));
            }
            else
            {
                viewModel.Update(e.Record);
            }

            var stale = otherList.FirstOrDefault(vm => string.Equals(vm.EventId, e.Record.EventId, StringComparison.OrdinalIgnoreCase));
            if (stale is not null)
            {
                otherList.Remove(stale);
            }

            OngoingEvents.Refresh();
            ArchivedEvents.Refresh();
        }

        private async Task HandleMailDropAsync()
        {
            var explorer = _services.OutlookApplication.ActiveExplorer();
            if (explorer?.Selection is null || explorer.Selection.Count == 0)
            {
                return;
            }

            var firstItem = explorer.Selection[1] as MailItem;
            if (firstItem is null)
            {
                return;
            }

            var participants = MailParticipantExtractor.Capture(firstItem);
            var preferredTemplate = _services.TemplatePreferences.GetPreferredTemplate(participants);

            await _services.EventRepository.CreateFromMailAsync(firstItem, preferredTemplate ?? "GEN", participants);
        }

        private async Task RefreshAllEventsAsync()
        {
            // Cancel previous refresh if any
            if (_refreshCts != null)
            {
                _refreshCts.Cancel();
                try 
                {
                    // Give it a moment to cancel
                    await Task.Delay(100); 
                } 
                catch {}
                _refreshCts.Dispose();
                _refreshCts = null;
            }

            _refreshCts = new System.Threading.CancellationTokenSource();
            var token = _refreshCts.Token;

            IsRefreshing = true;
            RefreshProgressValue = 0;
            RefreshStatusText = "Starting refresh...";

            try
            {
                var events = _services.EventRepository.GetAll().Where(e => e.Status == EventStatus.Open).ToList();
                var total = events.Count;
                var current = 0;

                foreach (var evt in events)
                {
                    if (token.IsCancellationRequested) break;

                    current++;
                    RefreshProgressValue = (double)current / total * 100;
                    RefreshStatusText = $"Refreshing {current}/{total}: {evt.EventTitle}";

                    // Trigger catch-up for each event
                    var conversationIds = evt.ConversationIds.ToList();
                    if (conversationIds.Count > 0)
                    {
                        // Now we await the task which completes when processing is done
                        await _services.EventMonitor.TriggerCatchUpAsync(evt.EventId, conversationIds, runImmediately: true);
                    }
                    
                    // Small delay to keep UI responsive and allow cancellation to be processed
                    await Task.Delay(10, token);
                }
                
                if (!token.IsCancellationRequested)
                {
                    RefreshStatusText = "Reloading view...";
                    await ReloadAsync();
                    RefreshStatusText = "Refresh completed.";
                    
                    // Keep the success message for a moment
                    await Task.Delay(2000, token);
                }
            }
            catch (OperationCanceledException)
            {
                RefreshStatusText = "Refresh cancelled.";
            }
            catch (System.Exception ex)
            {
                RefreshStatusText = $"Error: {ex.Message}";
                DebugLogger.Log($"RefreshAllEventsAsync error: {ex}");
            }
            finally
            {
                // Only reset IsRefreshing if this is the current task
                if (_refreshCts != null && _refreshCts.Token == token)
                {
                    IsRefreshing = false;
                    _refreshCts.Dispose();
                    _refreshCts = null;
                }
            }
        }

        private void OpenRuleManager()
        {
            var vm = new TemplateRuleManagerViewModel(_services);
            var view = new TemplateRuleManagerView { DataContext = vm };
            view.ShowDialog();
        }

        private void OpenTemplateEditor()
        {
            if (_openedEditor != null && _openedEditor.IsLoaded)
            {
                _openedEditor.Activate();
                if (_openedEditor.WindowState == WindowState.Minimized)
                {
                    _openedEditor.WindowState = WindowState.Normal;
                }
                return;
            }

            var vm = new TemplateEditorViewModel(_services);
            var view = new TemplateEditorView { DataContext = vm };

            var helper = new WindowInteropHelper(view);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            _openedEditor = view;
            view.Closed += (s, e) => _openedEditor = null;

            view.Show();
        }

        private async Task DeleteSelectionAsync()
        {
            if (SelectedEventIds.Count == 0)
            {
                return;
            }

            var message = $"Are you sure you want to delete {SelectedEventIds.Count} event(s)? This action cannot be undone.";
            if (System.Windows.MessageBox.Show(message, "Confirm Delete", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning) != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }

            IsBusy = true;
            try
            {
                var idsToDelete = SelectedEventIds.ToList();
                await _services.EventRepository.DeleteAsync(idsToDelete);
                SelectedEventIds.Clear();
                BatchArchiveCommand.RaiseCanExecuteChanged();
                DeleteCommand.RaiseCanExecuteChanged();
            }
            finally
            {
                IsBusy = false;
            }
        }

        public string RefreshStatusText
        {
            get => _refreshStatusText;
            private set
            {
                if (_refreshStatusText == value) return;
                _refreshStatusText = value;
                RaisePropertyChanged();
            }
        }

        public double RefreshProgressValue
        {
            get => _refreshProgressValue;
            private set
            {
                if (Math.Abs(_refreshProgressValue - value) < 0.001) return;
                _refreshProgressValue = value;
                RaisePropertyChanged();
            }
        }

        public bool IsRefreshing
        {
            get => _isRefreshing;
            private set
            {
                if (_isRefreshing == value) return;
                _isRefreshing = value;
                RaisePropertyChanged();
            }
        }

        private void LoadGlobalScripts()
        {
            GlobalScripts.Clear();
            foreach (var script in _services.PythonScripts.DiscoverScripts().Where(s => s.IsGlobal))
            {
                GlobalScripts.Add(script);
            }
        }
        
        private async Task ExecuteGlobalScriptAsync()
        {
            if (SelectedGlobalScript == null) return;
            
            try
            {
                IsBusy = true;
                
                var allEvents = _ongoingItems.Select(vm => _services.EventRepository.GetEvent(vm.EventId)).Where(e => e != null).ToList();
                
                var globalDataPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"osem-global-data-{Guid.NewGuid():N}.json");
                var globalDataJson = Newtonsoft.Json.JsonConvert.SerializeObject(allEvents, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(globalDataPath, globalDataJson);
                
                var context = new PythonScriptExecutionContext
                {
                    EventId = "GLOBAL",
                    EventTitle = "Global Execution",
                    DashboardValues = new Dictionary<string, string>
                    {
                        { "GlobalDataPath", globalDataPath },
                        { "IsGlobalExecution", "true" }
                    }
                };
                
                var exitCode = await _services.PythonScripts.ExecuteAsync(SelectedGlobalScript, context);
                
                if (exitCode == 0)
                {
                    MessageBox.Show("Global script executed successfully.");
                }
                else
                {
                    MessageBox.Show($"Global script failed with exit code {exitCode}.");
                }
                
                try { System.IO.File.Delete(globalDataPath); } catch {}
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error executing global script: {ex.Message}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        public async Task ExportEventsAsync(ExportOptionsViewModel options)
        {
            if (options.SelectedTemplate == null) return;

            IsBusy = true;
            try
            {
                var allRecords = await _services.EventRepository.GetAllAsync();
                var filtered = allRecords.Where(r =>
                {
                    bool templateMatch = string.Equals(r.DashboardTemplateId, options.SelectedTemplate.TemplateId, StringComparison.OrdinalIgnoreCase);
                    if (!templateMatch && string.IsNullOrEmpty(r.DashboardTemplateId) && string.Equals(options.SelectedTemplate.TemplateId, "GEN", StringComparison.OrdinalIgnoreCase))
                    {
                        templateMatch = true;
                    }

                    bool statusMatch = (options.ExportInProgress && r.Status == EventStatus.Open) ||
                                       (options.ExportArchived && r.Status == EventStatus.Archived);

                    return templateMatch &&
                           statusMatch &&
                           r.LastUpdatedOn.Date >= options.StartDate.Date &&
                           r.LastUpdatedOn.Date <= options.EndDate.Date;
                }).ToList();

                if (filtered.Count == 0)
                {
                    MessageBox.Show(Properties.Resources.No_matching_events_to_export, Properties.Resources.Export, MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                await Task.Run(async () =>
                {
                    // 1. Export Global CSV Report
                    if (options.ExportDashboardData)
                    {
                        string csvName = $"Report_{options.SelectedTemplate.DisplayName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                        foreach (var c in System.IO.Path.GetInvalidFileNameChars()) csvName = csvName.Replace(c, '_');
                        
                        await _services.CsvExport.ExportAsync(filtered, options.SelectedTemplate, options.TargetPath, csvName);
                    }

                    // 2. Export Attachments / Files (if requested)
                    if (options.ExportEventAttachments || options.ExportAdditionalFiles)
                    {
                        // Prepare allowed extensions set
                        HashSet<string>? allowedExtensions = null;
                        if (!options.ExportAllFileTypes)
                        {
                            allowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            foreach (var typeOption in options.FileTypeOptions)
                            {
                                if (typeOption.IsSelected)
                                {
                                    foreach (var ext in typeOption.Extensions)
                                    {
                                        allowedExtensions.Add(ext);
                                    }
                                }
                            }
                        }

                        foreach (var record in filtered)
                        {
                            string folderName = record.EventTitle;
                            if (options.FolderNamingMode.StartsWith("Key: "))
                            {
                                var key = options.FolderNamingMode.Substring(5);
                                var item = record.DashboardItems.FirstOrDefault(i => i.Key == key);
                                if (item != null && !string.IsNullOrWhiteSpace(item.Value))
                                {
                                    folderName = item.Value;
                                }
                            }

                            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                            {
                                folderName = folderName.Replace(c, '_');
                            }
                            folderName = folderName.Trim();
                            if (string.IsNullOrEmpty(folderName)) folderName = record.EventId;

                            var eventFolderPath = System.IO.Path.Combine(options.TargetPath, folderName);
                            System.IO.Directory.CreateDirectory(eventFolderPath);

                            if (options.ExportEventAttachments)
                            {
                                var attachmentsPath = System.IO.Path.Combine(eventFolderPath, "Attachments");
                                System.IO.Directory.CreateDirectory(attachmentsPath);

                                foreach (var att in record.Attachments)
                                {
                                    var ext = System.IO.Path.GetExtension(att.FileName).TrimStart('.').ToLowerInvariant();
                                    if ((options.ExportAllFileTypes || (allowedExtensions != null && allowedExtensions.Contains(ext))) && !string.IsNullOrEmpty(att.SourceMailEntryId))
                                    {
                                        await _dispatcher.InvokeAsync(() =>
                                        {
                                            try
                                            {
                                                var mailItem = _services.OutlookApplication.Session.GetItemFromID(att.SourceMailEntryId) as MailItem;
                                                if (mailItem != null)
                                                {
                                                    foreach (Attachment outlookAtt in mailItem.Attachments)
                                                    {
                                                        if (outlookAtt.FileName == att.FileName)
                                                        {
                                                            var savePath = System.IO.Path.Combine(attachmentsPath, att.FileName);
                                                            outlookAtt.SaveAsFile(savePath);
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            catch (System.Exception ex)
                                            {
                                                DebugLogger.Log($"Failed to export attachment {att.FileName}: {ex.Message}");
                                            }
                                        });
                                    }
                                }
                            }

                            if (options.ExportAdditionalFiles)
                            {
                                 var timePath = System.IO.Path.Combine(eventFolderPath, "AdditionalFiles");
                                 System.IO.Directory.CreateDirectory(timePath);

                                 // 1. Export User Added Files (Always export)
                                 if (record.AdditionalFiles != null)
                                 {
                                     foreach(var file in record.AdditionalFiles)
                                     {
                                         // Check if file exists at original path, if not try default storage (AppData)
                                         string sourceFilePath = file;
                                         if (!System.IO.File.Exists(sourceFilePath))
                                         {
                                             string storageRoot = OSEMAddIn.Properties.Settings.Default.EventFilesStoragePath;
                                             if (string.IsNullOrWhiteSpace(storageRoot))
                                             {
                                                 storageRoot = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "OSEMAddIn");
                                             }
                                             string localStorePath = System.IO.Path.Combine(storageRoot, "documents", record.EventId, System.IO.Path.GetFileName(file));
                                             if (System.IO.File.Exists(localStorePath))
                                             {
                                                 sourceFilePath = localStorePath;
                                             }
                                         }

                                         if(System.IO.File.Exists(sourceFilePath))
                                         {
                                             var fName = System.IO.Path.GetFileName(sourceFilePath);
                                             var ext = System.IO.Path.GetExtension(fName).TrimStart('.').ToLowerInvariant();
                                             
                                             if(options.ExportAllFileTypes || (allowedExtensions != null && allowedExtensions.Contains(ext)))
                                             {
                                                 try
                                                 {
                                                     System.IO.File.Copy(sourceFilePath, System.IO.Path.Combine(timePath, fName), true);
                                                 }
                                                 catch{}
                                             }
                                         }
                                     }
                                 }

                                 // 2. Export Template Files (Only if modified - Hash Check)
                                 if (options.SelectedTemplate != null && options.SelectedTemplate.AttachmentPaths != null)
                                 {
                                     foreach (var tmplFile in options.SelectedTemplate.AttachmentPaths)
                                     {
                                         if (record.ExcludedTemplateFiles != null && record.ExcludedTemplateFiles.Contains(tmplFile))
                                             continue;

                                         // Construct local path
                                         string storageRoot = OSEMAddIn.Properties.Settings.Default.EventFilesStoragePath;
                                         if (string.IsNullOrWhiteSpace(storageRoot))
                                         {
                                             storageRoot = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "OSEMAddIn");
                                         }
                                         string localStorePath = System.IO.Path.Combine(storageRoot, "documents", record.EventId, System.IO.Path.GetFileName(tmplFile));

                                         bool isModified = false;
                                         string? sourceToExport = null;

                                         if (System.IO.File.Exists(localStorePath))
                                         {
                                             // Local copy exists, check if it differs from original
                                             if (System.IO.File.Exists(tmplFile))
                                             {
                                                 try
                                                 {
                                                     string localHash = ComputeFileHash(localStorePath);
                                                     string originalHash = ComputeFileHash(tmplFile);
                                                     if (localHash != originalHash)
                                                     {
                                                         isModified = true;
                                                         sourceToExport = localStorePath;
                                                     }
                                                 }
                                                 catch
                                                 {
                                                     // If we can't read files to hash, assume modified to be safe
                                                     isModified = true;
                                                     sourceToExport = localStorePath;
                                                 }
                                             }
                                             else
                                             {
                                                 // Original missing, but we have local. Treat as modified/unique.
                                                 isModified = true;
                                                 sourceToExport = localStorePath;
                                             }
                                         }
                                         // If local doesn't exist, it's definitely not modified (user hasn't touched it), so skip.

                                         if (isModified && sourceToExport != null)
                                         {
                                             var fName = System.IO.Path.GetFileName(tmplFile);
                                             var ext = System.IO.Path.GetExtension(fName).TrimStart('.').ToLowerInvariant();
                                             
                                             if(options.ExportAllFileTypes || (allowedExtensions != null && allowedExtensions.Contains(ext)))
                                             {
                                                 try
                                                 {
                                                     System.IO.File.Copy(sourceToExport, System.IO.Path.Combine(timePath, fName), true);
                                                 }
                                                 catch (System.Exception ex)
                                                 {
                                                     DebugLogger.Log($"Failed to export template file {fName}: {ex.Message}");
                                                 }
                                             }
                                         }
                                     }
                                 }
                            }
                        }
                    }
                });

                MessageBox.Show(Properties.Resources.Export_completed, Properties.Resources.Export, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(string.Format(Properties.Resources.Export_failed_ex_Message, ex.Message), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private static string ComputeFileHash(string filePath)
        {
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                using (var stream = System.IO.File.OpenRead(filePath))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }
    }
}
