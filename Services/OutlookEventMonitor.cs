#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class OutlookEventMonitor : IDisposable
        {
            private const string InternetMessageIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            private const string ThreadIndexProperty = "http://schemas.microsoft.com/mapi/proptag/0x00710102";
            private const string ConversationIndexProperty = "http://schemas.microsoft.com/mapi/proptag/0x7F101102";
            private const string InReplyToProperty = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
            private const string ReferencesProperty = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";
            private const string TransportHeadersProperty = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            private const int CatchUpBatchSize = 20;
            private const int ThreadIndexPrefixBytes = 27; // First 27 bytes anchor the conversation root

            private static readonly Regex HeaderReferenceRegex = new Regex("(?im)^(?:References|In-Reply-To):\\s*(?<value>.+)$", RegexOptions.Compiled);
            private static readonly Regex MessageIdRegex = new Regex("<(?<id>[^>]+)>", RegexOptions.Compiled);
            // Updated regex to support:
            // 1. Korean (제목) and Japanese (件名) headers
            // 2. Multiline/Folded headers (lines starting with whitespace)
            // Note: Used [ \t] instead of \s for folding to avoid matching empty lines (since \s includes newlines)
            private static readonly Regex HistoricalSubjectRegex = new Regex(@"(?:Subject|主题|主旨|標題|제목|件名)\s*[:：]\s*(?<subject>.+(?:\r?\n[ \t]+.+)*)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Multiline);

            private readonly Outlook.Application _application;
            private readonly IEventRepository _eventRepository;
            private static readonly TimeSpan[] DeferredRetryDelays =
            {
                TimeSpan.FromSeconds(20),
                TimeSpan.FromMinutes(1),
                TimeSpan.FromMinutes(3),
                TimeSpan.FromMinutes(5)
            };
            private static readonly TimeSpan CatchUpLookbackWindow = TimeSpan.FromDays(14);
            private static readonly TimeSpan CatchUpFullHistoryWindow = TimeSpan.FromDays(3650);
            private static readonly TimeSpan CatchUpInterval = TimeSpan.FromMinutes(15);
            private static readonly TimeSpan CatchUpInitialDelay = TimeSpan.FromSeconds(10);
            private static bool s_conversationFilterSupported = true;
            private readonly ConcurrentDictionary<string, byte> _deferredEntryTracker = new(StringComparer.OrdinalIgnoreCase);
            private readonly System.Threading.SemaphoreSlim _catchUpSemaphore = new(1, 1);
            private readonly ConcurrentQueue<(string EventId, string ConversationId)> _catchUpQueue = new();
            private readonly ConcurrentDictionary<string, byte> _catchUpTracker = new(StringComparer.OrdinalIgnoreCase);

            // Sync-Aware Search Fields
            private readonly ConcurrentQueue<string> _pendingSearchQueue = new();
            private readonly ConcurrentQueue<string> _pendingConversationSearchQueue = new();
            private Timer? _searchDebounceTimer;
            private Timer? _syncPollingTimer;
            private int _activeSyncCount;
            private Outlook.SyncObjects? _syncObjects;

            private const int SearchDebounceMs = 2000;
            private const int SyncPollingIntervalMs = 30000;
            // Increased lookback window to 1 hour to avoid missing items during long syncs or indexing delays
            private const int SearchLookbackMinutes = 60;
            private const string SyncRetryMarker = "OSEM_SYNC_RETRY_MARKER";

            // Retry Logic Fields
            private readonly ConcurrentDictionary<string, int> _searchRetryTracker = new(StringComparer.OrdinalIgnoreCase);
            private readonly ConcurrentDictionary<string, List<string>> _activeSearchMap = new(StringComparer.OrdinalIgnoreCase);
            private const int MaxSearchRetries = 10;

            public event EventHandler<string>? SearchStatusChanged;

            private Outlook.ApplicationEvents_11_Event? _applicationEvents;
            private Outlook.Items? _sentItemsEvents;
            private Outlook.Items? _inboxItemsEvents;
            private Outlook.NameSpace? _session;
            private Timer? _catchUpTimer;
            private bool _catchUpTimerInitialized;
            private bool _catchUpPausedBySync;
            private bool _isStarted;
            private bool _isDisposed;

            private readonly List<Outlook.Items> _customFolderMonitors = new List<Outlook.Items>();

            public OutlookEventMonitor(Outlook.Application application, IEventRepository eventRepository)
            {
                _application = application ?? throw new ArgumentNullException(nameof(application));
                _eventRepository = eventRepository ?? throw new ArgumentNullException(nameof(eventRepository));
            }

            public void Start()
            {
                ThrowIfDisposed();

                if (_isStarted)
                {
                    return;
                }

                _applicationEvents = (Outlook.ApplicationEvents_11_Event)_application;
                _applicationEvents.NewMailEx += OnNewMailEx;
                _application.AdvancedSearchComplete += OnAdvancedSearchComplete;

                try
                {
                    _syncObjects = _application.Session.SyncObjects;
                    for (int i = 1; i <= _syncObjects.Count; i++)
                    {
                        var syncObj = _syncObjects[i];
                        syncObj.SyncStart += OnSyncStart;
                        syncObj.SyncEnd += OnSyncEnd;
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Failed to setup sync monitoring: {ex.Message}");
                }

                _searchDebounceTimer = new Timer { Interval = SearchDebounceMs };
                _searchDebounceTimer.Tick += OnSearchDebounceTick;

                _syncPollingTimer = new Timer { Interval = SyncPollingIntervalMs };
                _syncPollingTimer.Tick += OnSyncPollingTick;
                _syncPollingTimer.Start();

                // Monitor Sent Items
                try
                {
                    var session = EnsureSession();
                    if (session != null)
                    {
                        var sentFolder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                        if (sentFolder != null)
                        {
                            _sentItemsEvents = sentFolder.Items;
                            _sentItemsEvents.ItemAdd += OnSentItemAdded;
                        }

                        var inboxFolder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        if (inboxFolder != null)
                        {
                            _inboxItemsEvents = inboxFolder.Items;
                            _inboxItemsEvents.ItemAdd += OnInboxItemAdded;
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Failed to subscribe to Sent/Inbox Items: {ex.Message}");
                }

                _catchUpTimer = new Timer
                {
                    Interval = Math.Max(1, (int)CatchUpInitialDelay.TotalMilliseconds)
                };
                _catchUpTimer.Tick += OnCatchUpTimerTick;
                _catchUpTimer.Start();

                _ = PopulateCatchUpQueueAsync();

                InitializeCustomMonitors();

                _isStarted = true;
            }

            public void Stop()
            {
                if (!_isStarted)

                {
                    return;
                }

                if (_applicationEvents is not null)
                {
                    _applicationEvents.NewMailEx -= OnNewMailEx;
                }

                try
                {
                    _application.AdvancedSearchComplete -= OnAdvancedSearchComplete;
                }
                catch { }

                if (_searchDebounceTimer != null)
                {
                    _searchDebounceTimer.Stop();
                    _searchDebounceTimer.Dispose();
                    _searchDebounceTimer = null;
                }

                if (_syncPollingTimer != null)
                {
                    _syncPollingTimer.Stop();
                    _syncPollingTimer.Dispose();
                    _syncPollingTimer = null;
                }

                if (_sentItemsEvents is not null)
                {
                    _sentItemsEvents.ItemAdd -= OnSentItemAdded;
                    Marshal.ReleaseComObject(_sentItemsEvents);
                    _sentItemsEvents = null;
                }

                if (_inboxItemsEvents is not null)
                {
                    _inboxItemsEvents.ItemAdd -= OnInboxItemAdded;
                    Marshal.ReleaseComObject(_inboxItemsEvents);
                    _inboxItemsEvents = null;
                }

                if (_catchUpTimer is not null)
                {
                    _catchUpTimer.Stop();
                    _catchUpTimer.Tick -= OnCatchUpTimerTick;
                    _catchUpTimer.Dispose();
                    _catchUpTimer = null;
                }

                // Clean up custom monitors
                foreach (var items in _customFolderMonitors)
                {
                    items.ItemAdd -= OnCustomFolderItemAdded;
                    Marshal.ReleaseComObject(items);
                }
                _customFolderMonitors.Clear();

                DrainCatchUpQueue();
                _catchUpTracker.Clear();
                _catchUpTimerInitialized = false;

                ReleaseComObject(ref _session);

                _applicationEvents = null;
                _isStarted = false;
            }

            public Task TriggerCatchUpAsync(string eventId, IEnumerable<string> conversationIds, bool runImmediately = false, TimeSpan? immediateTimeout = null, bool useFullHistory = false)
            {
                ThrowIfDisposed();

                if (!_isStarted)
                {
                    return Task.CompletedTask;
                }

                if (string.IsNullOrWhiteSpace(eventId) || conversationIds is null)
                {
                    return Task.CompletedTask;
                }

                var added = 0;
                var requested = 0;
                foreach (var conversationId in conversationIds)
                {
                    var normalized = conversationId?.Trim();
                    if (string.IsNullOrEmpty(normalized))
                    {
                        continue;
                    }

                    requested++;
                    var safeConversationId = normalized!;
                    var key = BuildCatchUpKey(eventId, safeConversationId);
                    if (_catchUpTracker.TryAdd(key, 0))
                    {
                        _catchUpQueue.Enqueue((eventId, safeConversationId));
                        added++;
                    }
                }

                var hasQueuedWork = !_catchUpQueue.IsEmpty;

                if (requested == 0 && !hasQueuedWork)
                {
                    DebugLogger.Log($"TriggerCatchUpAsync received no valid conversations for Event='{eventId}'.");
                    return Task.CompletedTask;
                }

                if (added > 0)
                {
                    DebugLogger.Log($"Catch-up request enqueued {added} conversations for Event='{eventId}'. Immediate={runImmediately} FullHistory={useFullHistory}");
                }
                else if (runImmediately && hasQueuedWork)
                {
                    DebugLogger.Log($"Catch-up request reused {(_catchUpQueue.Count > 0 ? _catchUpQueue.Count : requested)} queued conversations for Event='{eventId}'. Immediate={runImmediately} FullHistory={useFullHistory}");
                }
                else
                {
                    DebugLogger.Log($"Catch-up request had no new conversations for Event='{eventId}'. Immediate={runImmediately} FullHistory={useFullHistory}");
                }

                if (runImmediately)
                {
                    var remaining = added;
                    if (remaining == 0 && requested > 0)
                    {
                        remaining = requested;
                    }

                    if (remaining == 0 && hasQueuedWork)
                    {
                        remaining = Math.Max(1, _catchUpQueue.Count);
                    }

                    var timeout = immediateTimeout.HasValue && immediateTimeout.Value > TimeSpan.Zero
                        ? immediateTimeout.Value
                        : TimeSpan.Zero;
                    return StartImmediateCatchUpProcessing(eventId, remaining, timeout, useFullHistory);
                }

                return Task.CompletedTask;
            }
            private Task StartImmediateCatchUpProcessing(string eventId, int remaining, TimeSpan timeout, bool useFullHistory)
            {
                if (remaining <= 0)
                {
                    return Task.CompletedTask;
                }

                DebugLogger.Log($"TriggerCatchUpAsync immediate processing scheduled for Event='{eventId}' Remaining={remaining} Timeout={timeout.TotalSeconds:F1}s FullHistory={useFullHistory}");

                return Task.Run(async () =>
                {
                    try
                    {
                        var deadline = timeout > TimeSpan.Zero ? DateTime.UtcNow + timeout : DateTime.MaxValue;

                        while (remaining > 0)
                        {
                            int? waitTimeoutMs = null;
                            if (timeout > TimeSpan.Zero)
                            {
                                var remainingMs = (int)Math.Ceiling((deadline - DateTime.UtcNow).TotalMilliseconds);
                                if (remainingMs <= 0)
                                {
                                    DebugLogger.Log($"TriggerCatchUpAsync immediate processing timeout reached for Event='{eventId}' Remaining={remaining}");
                                    break;
                                }

                                waitTimeoutMs = Math.Min(remainingMs, int.MaxValue);
                            }

                            var processed = await ProcessCatchUpBatchAsync(remaining, waitForLock: true, lockTimeoutMilliseconds: waitTimeoutMs, useFullHistory: useFullHistory, preferEventId: eventId).ConfigureAwait(false);
                            if (processed == 0)
                            {
                                if (timeout == TimeSpan.Zero)
                                {
                                    DebugLogger.Log($"TriggerCatchUpAsync immediate processing paused for Event='{eventId}' (processed=0).");
                                    break;
                                }

                                await Task.Delay(200).ConfigureAwait(false);
                                continue;
                            }

                            remaining = Math.Max(0, remaining - processed);
                            DebugLogger.Log($"TriggerCatchUpAsync immediate processing progress for Event='{eventId}' Remaining={remaining}");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"TriggerCatchUpAsync immediate processing exception for Event='{eventId}': {ex.Message}");
                    }
                });
            }

            public void Dispose()
            {
                if (_isDisposed)
                {
                    return;
                }

                Stop();
                _catchUpSemaphore.Dispose();
                _isDisposed = true;
            }

            private void OnNewMailEx(string entryIdCollection)
            {
                if (!_isStarted || string.IsNullOrWhiteSpace(entryIdCollection))
                {
                    return;
                }

                DebugLogger.Log($"OnNewMailEx received ids='{entryIdCollection}'");
                _ = Task.Run(() => ProcessEntryIdsAsync(entryIdCollection));
            }

            private void OnSyncStart()
            {
                var count = System.Threading.Interlocked.Increment(ref _activeSyncCount);
                if (count == 1)
                {
                    SearchStatusChanged?.Invoke(this, "Waiting for Outlook sync...");
                }
            }

            private void OnSyncEnd()
            {
                var count = System.Threading.Interlocked.Decrement(ref _activeSyncCount);
                if (count < 0)
                {
                    System.Threading.Interlocked.Exchange(ref _activeSyncCount, 0);
                    count = 0;
                }

                if (count == 0)
                {
                    SearchStatusChanged?.Invoke(this, "Ready");
                }

                if (_pendingSearchQueue.Count > 0 || _pendingConversationSearchQueue.Count > 0)
                {
                    DebugLogger.Log("SyncEnd detected. Triggering search debounce.");
                    _searchDebounceTimer?.Stop();
                    _searchDebounceTimer?.Start();
                }

                if (_catchUpPausedBySync)
                {
                    DebugLogger.Log("SyncEnd detected. Resuming paused catch-up.");
                    _catchUpPausedBySync = false;
                    OnCatchUpTimerTick(this, EventArgs.Empty);
                }
            }

            private void OnSyncPollingTick(object? sender, EventArgs e)
            {
                var hasItems = !_pendingSearchQueue.IsEmpty || !_pendingConversationSearchQueue.IsEmpty;
                
                if (hasItems)
                {
                    if (!IsOutlookSyncing())
                    {
                        DebugLogger.Log("SyncPollingTick: Queue has items and Outlook is idle. Triggering search.");
                        _searchDebounceTimer?.Stop();
                        _searchDebounceTimer?.Start();
                    }
                    else
                    {
                        // Modified: Wait for sync to complete instead of forcing search to avoid UI hang.
                        // The search will be triggered by OnSyncEnd or the next polling tick when sync is done.
                        DebugLogger.Log("SyncPollingTick: Queue has items but Outlook is syncing. Waiting for sync to complete.");
                    }
                }
            }

            private bool IsOutlookSyncing()
            {
                return _activeSyncCount > 0;
            }

            private void OnSearchDebounceTick(object? sender, EventArgs e)
            {
                _searchDebounceTimer?.Stop();

                if (IsOutlookSyncing())
                {
                    DebugLogger.Log("SearchDebounce: Outlook still syncing. Rescheduling.");
                    _searchDebounceTimer?.Start();
                    return;
                }

                PerformAdvancedSearch();
            }

            private void PerformAdvancedSearch()
            {
                try
                {
                    var hasEntryIds = !_pendingSearchQueue.IsEmpty;
                    var hasConversations = !_pendingConversationSearchQueue.IsEmpty;

                    if (!hasEntryIds && !hasConversations)
                    {
                        return;
                    }

                    // Capture EntryIDs instead of discarding them, for logging purposes
                    var targetEntryIds = new List<string>();
                    var isRetry = false;
                    while (_pendingSearchQueue.TryDequeue(out var eid)) 
                    { 
                        if (eid == SyncRetryMarker)
                        {
                            isRetry = true;
                        }
                        else
                        {
                            targetEntryIds.Add(eid);
                        }
                    }
                    
                    // If we are processing real items while Outlook is syncing, schedule a retry for later
                    // to ensure we catch items that might be missed due to indexing lag.
                    if (targetEntryIds.Count > 0 && IsOutlookSyncing())
                    {
                        DebugLogger.Log("PerformAdvancedSearch: Sync active during search. Re-queuing items for later.");
                        foreach (var id in targetEntryIds) _pendingSearchQueue.Enqueue(id);
                        if (isRetry) _pendingSearchQueue.Enqueue(SyncRetryMarker);
                        return;
                    }

                    var conversations = new HashSet<string>();
                    while (_pendingConversationSearchQueue.TryDequeue(out var cid))
                    {
                        conversations.Add(cid);
                    }

                    if (SearchStatusChanged != null)
                    {
                        SearchStatusChanged(this, "Searching...");
                    }

                    var app = Globals.ThisAddIn.Application;
                    var filters = new List<string>();

                    if (hasEntryIds || isRetry)
                    {
                        // Use the expanded lookback window
                        var timeThreshold = DateTime.Now.AddMinutes(-SearchLookbackMinutes);
                        // Use PR_CREATION_TIME (0x30070040) instead of datereceived to catch items synced recently
                        // regardless of when they were originally sent/received by the server.
                        filters.Add($"\"http://schemas.microsoft.com/mapi/proptag/0x30070040\" > '{timeThreshold:g}'");
                        if (targetEntryIds.Count > 0)
                        {
                            DebugLogger.Log($"PerformAdvancedSearch targeting {targetEntryIds.Count} pending EntryIDs with creation time threshold: {timeThreshold:g}");
                        }
                        else
                        {
                            DebugLogger.Log($"PerformAdvancedSearch executing retry/catch-up with creation time threshold: {timeThreshold:g}");
                        }
                    }

                    foreach (var cid in conversations)
                    {
                         filters.Add($"\"http://schemas.microsoft.com/mapi/proptag/0x30130102\" = '{cid}'");
                    }

                    if (filters.Count == 0) return;

                    string filter = string.Join(" OR ", filters);
                    
                    // Generate a unique tag for this search to track retries
                    var searchTag = $"OSEM_Recovery_Search_{Guid.NewGuid()}";
                    if (targetEntryIds.Count > 0)
                    {
                        _activeSearchMap[searchTag] = targetEntryIds;
                    }

                    DebugLogger.Log($"Starting AdvancedSearch with filter: {filter} Tag='{searchTag}'");
                    
                    Outlook.NameSpace ns = app.GetNamespace("MAPI");
                    Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Store store = inbox.Store;
                    string scopePath = "'" + store.GetRootFolder().FolderPath + "'";
                    
                    app.AdvancedSearch(scopePath, filter, true, searchTag);
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"PerformAdvancedSearch failed: {ex.Message}");
                    if (SearchStatusChanged != null) SearchStatusChanged(this, "Ready");
                }
            }

            private void OnAdvancedSearchComplete(Outlook.Search searchObject)
            {
                if (!searchObject.Tag.StartsWith("OSEM_Recovery_Search"))
                {
                    return;
                }

                DebugLogger.Log($"AdvancedSearch completed (Tag={searchObject.Tag}). Found {searchObject.Results.Count} items.");
                
                if (SearchStatusChanged != null)
                {
                    SearchStatusChanged(this, "Ready");
                }

                // Retry Logic
                if (searchObject.Results.Count == 0)
                {
                    if (_activeSearchMap.TryRemove(searchObject.Tag, out var searchedIds) && searchedIds != null && searchedIds.Count > 0)
                    {
                        var toRetry = new List<string>();
                        foreach (var id in searchedIds)
                        {
                            var count = _searchRetryTracker.AddOrUpdate(id, 1, (k, v) => v + 1);
                            if (count <= MaxSearchRetries)
                            {
                                toRetry.Add(id);
                            }
                            else
                            {
                                DebugLogger.Log($"Search retry exhausted for EntryID {id}");
                                _searchRetryTracker.TryRemove(id, out _);
                            }
                        }

                        if (toRetry.Count > 0)
                        {
                            DebugLogger.Log($"AdvancedSearch found 0 items. Retrying {toRetry.Count} EntryIDs (Attempt {(toRetry.Count > 0 ? _searchRetryTracker[toRetry[0]] : 0)}).");
                            
                            // Delay the retry slightly to avoid hammering
                            Task.Delay(5000).ContinueWith(_ => 
                            {
                                foreach (var id in toRetry) _pendingSearchQueue.Enqueue(id);
                            });
                        }
                    }
                }
                else
                {
                    // Found items! Clear retry tracker for this batch (heuristic)
                    if (_activeSearchMap.TryRemove(searchObject.Tag, out var searchedIds))
                    {
                        foreach (var id in searchedIds) _searchRetryTracker.TryRemove(id, out _);
                    }
                }

                foreach (var item in searchObject.Results)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        try
                        {
                            Outlook.MAPIFolder? parentFolder = null;
                            var snapshot = BuildSnapshot(mailItem, out parentFolder);
                            
                            if (snapshot != null)
                            {
                                _ = _eventRepository.TryAddMailAsync(snapshot);
                            }
                            
                            if (parentFolder != null)
                            {
                                Marshal.ReleaseComObject(parentFolder);
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"Error processing search result: {ex.Message}");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mailItem);
                        }
                    }
                }
            }

            private void OnSentItemAdded(object item)
            {
                if (!_isStarted || item is not Outlook.MailItem mailItem)
                {
                    return;
                }

                try
                {
                    var entryId = mailItem.EntryID;
                    if (string.IsNullOrWhiteSpace(entryId))
                    {
                        return;
                    }

                    DebugLogger.Log($"OnSentItemAdded received EntryId='{entryId}' Subject='{mailItem.Subject}'");
                    _ = Task.Run(() => ProcessEntryIdsAsync(entryId));
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"OnSentItemAdded failed: {ex.Message}");
                }
                finally
                {
                    // Note: We do not release mailItem here because it is passed by the event and might be used elsewhere?
                    // Actually, for ItemAdd event, we should be careful. But usually it's safe to release if we are done.
                    // However, since we are passing EntryID to async task, we don't need the object anymore.
                    Marshal.ReleaseComObject(mailItem);
                }
            }

            private void OnInboxItemAdded(object item)
            {
                if (!_isStarted) return;

                if (item is Outlook.MailItem mailItem)
                {
                    try
                    {
                        var entryId = mailItem.EntryID;
                        if (!string.IsNullOrWhiteSpace(entryId))
                        {
                            DebugLogger.Log($"OnInboxItemAdded received EntryId='{entryId}' Subject='{mailItem.Subject}'");
                            _ = Task.Run(() => ProcessEntryIdsAsync(entryId));
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"OnInboxItemAdded failed: {ex.Message}");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(mailItem);
                    }
                }
                else if (item != null && Marshal.IsComObject(item))
                {
                    Marshal.ReleaseComObject(item);
                }
            }

            private async void OnCatchUpTimerTick(object? sender, EventArgs e)
            {
                if (!_isStarted)
                {
                    return;
                }

                if (IsOutlookSyncing())
                {
                    DebugLogger.Log("CatchUpTimerTick: Outlook is syncing. Pausing catch-up until sync completes.");
                    _catchUpPausedBySync = true;
                    return;
                }

                if (!_catchUpTimerInitialized && _catchUpTimer is not null)
                {
                    _catchUpTimerInitialized = true;
                    _catchUpTimer.Interval = Math.Max(1, (int)CatchUpInterval.TotalMilliseconds);
                }

                _ = await Task.Run(async () =>
                {
                    try
                    {
                        return await ProcessCatchUpBatchAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Background catch-up batch failed: {ex.Message}");
                        return 0;
                    }
                }).ConfigureAwait(false);
            }

            private async Task ProcessEntryIdsAsync(string entryIdCollection)
            {
                var entryIds = entryIdCollection
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(id => id.Trim())
                    .Where(id => id.Length > 0)
                    .ToList();

                if (entryIds.Count == 0)
                {
                    return;
                }

                var session = EnsureSession();
                if (session is null)
                {
                    DebugLogger.Log("ProcessEntryIdsAsync deferred because Outlook session is unavailable.");
                    foreach (var entryId in entryIds)
                    {
                        ScheduleDeferredRetry(entryId, 0);
                    }

                    return;
                }

                foreach (var entryId in entryIds)
                {
                    var processed = await TryProcessEntryAsync(session, entryId, isDeferred: false, delayIndex: 0).ConfigureAwait(false);
                    if (!processed)
                    {
                        DebugLogger.Log($"EntryID {entryId} failed to resolve. Queuing for Advanced Search.");
                        _pendingSearchQueue.Enqueue(entryId);
                    }
                }
            }

            private async Task<bool> TryProcessEntryAsync(Outlook.NameSpace session, string entryId, bool isDeferred, int delayIndex)
            {
                Outlook.MailItem? mailItem = null;
                Outlook.MAPIFolder? parentFolder = null;

                try
                {
                    mailItem = await ResolveMailItemWithRetryAsync(session, entryId).ConfigureAwait(false);
                    if (mailItem is null)
                    {
                        DebugLogger.Log($"TryProcessEntryAsync {(isDeferred ? "deferred" : "initial")} attempt could not resolve entry '{entryId}' (delayIndex={delayIndex})");
                        return false;
                    }

                    var snapshot = BuildSnapshot(mailItem, out parentFolder);
                    if (snapshot is null)
                    {
                        DebugLogger.Log($"TryProcessEntryAsync {(isDeferred ? "deferred" : "initial")} attempt produced no snapshot for entry '{entryId}' (delayIndex={delayIndex})");
                        return false;
                    }

                    DebugLogger.Log($"TryProcessEntryAsync captured Entry='{snapshot.EntryId}' Store='{snapshot.StoreId}' Conv='{snapshot.ConversationId}' MsgId='{snapshot.InternetMessageId}' Attachments={snapshot.Attachments.Count}");
                    await _eventRepository.TryAddMailAsync(snapshot).ConfigureAwait(false);
                    return true;
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine($"[OutlookEventMonitor] Failed to process entry '{entryId}': {ex}");
                    DebugLogger.Log($"TryProcessEntryAsync exception for entry '{entryId}': {ex.Message}");
                    return false;
                }
                finally
                {
                    if (parentFolder is not null)
                    {
                        Marshal.ReleaseComObject(parentFolder);
                        parentFolder = null;
                    }

                    if (mailItem is not null)
                    {
                        Marshal.ReleaseComObject(mailItem);
                        mailItem = null;
                    }
                }
            }

            private void ScheduleDeferredRetry(string entryId, int delayIndex)
            {
                if (delayIndex >= DeferredRetryDelays.Length)
                {
                    DebugLogger.Log($"Deferred processing exhausted retries for entry '{entryId}'.");
                    return;
                }

                if (!_deferredEntryTracker.TryAdd(entryId, 0))
                {
                    return;
                }

                var delay = DeferredRetryDelays[delayIndex];
                DebugLogger.Log($"Scheduling deferred retry {delayIndex + 1} for entry '{entryId}' in {delay.TotalSeconds} seconds.");

                _ = Task.Run(async () =>
                {
                    var shouldContinue = false;
                    try
                    {
                        await Task.Delay(delay).ConfigureAwait(false);
                        shouldContinue = await ProcessDeferredEntryAsync(entryId, delayIndex).ConfigureAwait(false);
                    }
                    finally
                    {
                        _deferredEntryTracker.TryRemove(entryId, out _);
                    }

                    if (shouldContinue)
                    {
                        ScheduleDeferredRetry(entryId, delayIndex + 1);
                    }
                });
            }

            private async Task<bool> ProcessDeferredEntryAsync(string entryId, int delayIndex)
            {
                var session = EnsureSession();
                if (session is null)
                {
                    DebugLogger.Log($"Deferred retry {delayIndex + 1} skipped for entry '{entryId}' because Outlook session is unavailable.");
                    return true;
                }

                var processed = await TryProcessEntryAsync(session, entryId, isDeferred: true, delayIndex: delayIndex).ConfigureAwait(false);
                return !processed;
            }

            private Outlook.MailItem? ResolveMailItem(Outlook.NameSpace session, string entryId)
            {
                try
                {
                    var item = session.GetItemFromID(entryId) as Outlook.MailItem;
                    if (item is not null)
                    {
                        DebugLogger.Log($"ResolveMailItem default store hit for entry '{entryId}'");
                        return item;
                    }
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"ResolveMailItem default store miss for entry '{entryId}', hresult=0x{ex.HResult:X8}");
                }

                return null;
            }

            private async Task<int> ProcessCatchUpBatchAsync(int maxBatch = CatchUpBatchSize, bool waitForLock = false, int? lockTimeoutMilliseconds = null, bool useFullHistory = false, string? preferEventId = null)
            {
                if (!_isStarted)
                {
                    return 0;
                }

                DebugLogger.Log($"ProcessCatchUpBatchAsync invoked maxBatch={maxBatch} waitForLock={waitForLock} lockTimeout={lockTimeoutMilliseconds?.ToString() ?? "null"} fullHistory={useFullHistory} preferEvent='{preferEventId ?? string.Empty}'");

                var limit = Math.Max(1, maxBatch);
                var processed = 0;
                var metadataCache = new Dictionary<string, CatchUpMetadata?>(StringComparer.OrdinalIgnoreCase);
                var populateAttempted = false;

                try
                {
                    while (processed < limit)
                    {
                        (string EventId, string ConversationId)? request = null;
                        var acquired = false;

                        try
                        {
                            if (waitForLock)
                            {
                                if (lockTimeoutMilliseconds.HasValue)
                                {
                                    var timeout = Math.Max(0, lockTimeoutMilliseconds.Value);
                                    acquired = await _catchUpSemaphore.WaitAsync(timeout).ConfigureAwait(false);
                                    if (!acquired)
                                    {
                                        DebugLogger.Log("ProcessCatchUpBatchAsync semaphore wait timed out.");
                                        break;
                                    }
                                }
                                else
                                {
                                    await _catchUpSemaphore.WaitAsync().ConfigureAwait(false);
                                    acquired = true;
                                }
                            }
                            else
                            {
                                acquired = await _catchUpSemaphore.WaitAsync(0).ConfigureAwait(false);
                                if (!acquired)
                                {
                                    DebugLogger.Log("ProcessCatchUpBatchAsync semaphore unavailable.");
                                    break;
                                }
                            }

                            if (_catchUpQueue.IsEmpty)
                            {
                                request = null;
                            }
                            else if (!string.IsNullOrWhiteSpace(preferEventId))
                            {
                                var queueCount = _catchUpQueue.Count;
                                var rotated = 0;
                                while (rotated++ < queueCount && _catchUpQueue.TryDequeue(out var candidate))
                                {
                                    if (string.Equals(candidate.EventId, preferEventId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        request = candidate;
                                        break;
                                    }

                                    _catchUpQueue.Enqueue(candidate);
                                }

                                if (!request.HasValue && rotated > 1)
                                {
                                    DebugLogger.Log($"ProcessCatchUpBatchAsync preferEvent='{preferEventId}' not found in queue during scan.");
                                }
                            }

                            if (!request.HasValue && !_catchUpQueue.IsEmpty)
                            {
                                if (_catchUpQueue.TryDequeue(out var candidate))
                                {
                                    request = candidate;
                                }
                            }

                            if (request.HasValue)
                            {
                                var key = BuildCatchUpKey(request.Value.EventId, request.Value.ConversationId);
                                _catchUpTracker.TryRemove(key, out _);
                            }
                        }
                        finally
                        {
                            if (acquired)
                            {
                                _catchUpSemaphore.Release();
                            }
                        }

                        if (!request.HasValue)
                        {
                            if (!populateAttempted)
                            {
                                populateAttempted = true;
                                DebugLogger.Log("ProcessCatchUpBatchAsync queue empty; populating from repository.");
                                await PopulateCatchUpQueueAsync().ConfigureAwait(false);
                                if (!_catchUpQueue.IsEmpty)
                                {
                                    continue;
                                }

                                DebugLogger.Log("ProcessCatchUpBatchAsync queue still empty after populate.");
                            }

                            break;
                        }

                        var requestValue = request.Value;
                        var session = EnsureSession();
                        if (session is null)
                        {
                            DebugLogger.Log("Catch-up tick skipped because Outlook session is unavailable.");
                            EnqueueCatchUpRequest(requestValue);

                            if (!waitForLock)
                            {
                                break;
                            }

                            await Task.Delay(200).ConfigureAwait(false);
                            continue;
                        }

                        if (!metadataCache.TryGetValue(requestValue.EventId, out var metadata))
                        {
                            DebugLogger.Log($"ProcessCatchUpBatchAsync loading catch-up metadata for Event='{requestValue.EventId}'");
                            metadata = await GetCatchUpMetadataAsync(requestValue.EventId).ConfigureAwait(false);
                            metadataCache[requestValue.EventId] = metadata;
                        }

                        DebugLogger.Log($"ProcessCatchUpBatchAsync processing Event='{requestValue.EventId}' Conversation='{requestValue.ConversationId}'");
                        await CatchUpConversationAsync(session, requestValue.ConversationId, requestValue.EventId, metadata, useFullHistory).ConfigureAwait(false);
                        DebugLogger.Log($"ProcessCatchUpBatchAsync completed Event='{requestValue.EventId}' Conversation='{requestValue.ConversationId}'");

                        processed++;
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Catch-up batch exception: {ex.Message}");
                }

                DebugLogger.Log($"ProcessCatchUpBatchAsync exiting processed={processed} preferEvent='{preferEventId ?? string.Empty}'");
                return processed;
            }

            private async Task PopulateCatchUpQueueAsync()
            {
                try
                {
                    var events = await _eventRepository.GetAllAsync().ConfigureAwait(false);
                    var openEvents = events
                        .Where(record => record.Status == EventStatus.Open && record.ConversationIds.Count > 0)
                        .ToList();

                    if (openEvents.Count == 0)
                    {
                        DebugLogger.Log("Catch-up queue population found no open events with tracked conversations.");
                        return;
                    }

                    var enqueued = 0;
                    foreach (var record in openEvents)
                    {
                        foreach (var conversationId in record.ConversationIds
                                     .Where(id => !string.IsNullOrWhiteSpace(id))
                                     .Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            var key = BuildCatchUpKey(record.EventId, conversationId);
                            if (_catchUpTracker.TryAdd(key, 0))
                            {
                                _catchUpQueue.Enqueue((record.EventId, conversationId));
                                enqueued++;
                            }
                        }
                    }

                    if (enqueued > 0)
                    {
                        DebugLogger.Log($"Catch-up queue populated with {enqueued} conversations.");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Catch-up queue population failed: {ex.Message}");
                }
            }

            private void EnqueueCatchUpRequest((string EventId, string ConversationId) request)
            {
                var key = BuildCatchUpKey(request.EventId, request.ConversationId);
                if (_catchUpTracker.TryAdd(key, 0))
                {
                    _catchUpQueue.Enqueue(request);
                }
            }

            private static string BuildCatchUpKey(string eventId, string conversationId)
            {
                return string.Concat(eventId ?? string.Empty, "::", conversationId ?? string.Empty);
            }

            private async Task CatchUpConversationAsync(Outlook.NameSpace session, string conversationId, string eventId, CatchUpMetadata? metadata, bool useFullHistory)
            {
                var lookbackWindow = useFullHistory ? CatchUpFullHistoryWindow : CatchUpLookbackWindow;
                var cutoffUtc = DateTime.UtcNow.Subtract(lookbackWindow);

                if (metadata?.EarliestReceivedOnUtc is DateTime earliestTrackedUtc)
                {
                    var extendedCutoff = earliestTrackedUtc.AddHours(-12);
                    if (extendedCutoff < cutoffUtc)
                    {
                        cutoffUtc = extendedCutoff;
                        DebugLogger.Log($"CatchUpConversationAsync extended lookback for Event='{eventId}' Conversation='{conversationId}' to {cutoffUtc:O}");
                    }
                }

                DebugLogger.Log($"CatchUpConversationAsync scanning Event='{eventId}' Conversation='{conversationId}' FullHistory={useFullHistory} LookbackDays={lookbackWindow.TotalDays:F1}");
                // Recursive Search Loop Start
                IReadOnlyList<(string EntryId, string StoreId)> references;
                try
                {
                    references = CollectConversationMailReferences(session, eventId, conversationId, cutoffUtc, metadata);
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"CollectConversationMailReferences failed for Event='{eventId}' Conversation='{conversationId}': {ex}");
                    references = Array.Empty<(string EntryId, string StoreId)>(); // Verified
                }
                if (references.Count == 0)
                {
                    DebugLogger.Log($"CatchUpConversationAsync no references found Event='{eventId}' Conversation='{conversationId}' FullHistory={useFullHistory}");
                    return;
                }

                DebugLogger.Log($"CatchUpConversationAsync found {references.Count} references Event='{eventId}' Conversation='{conversationId}' FullHistory={useFullHistory}");
                // Loop Start
                foreach (var reference in references)
                {
                    await Task.Yield();

                    Outlook.MailItem? mailItem = null;
                    Outlook.MAPIFolder? parentFolder = null;

                    try
                    {
                        mailItem = ResolveMailItem(session, reference.EntryId);
                        if (mailItem is null)
                        {
                            _pendingConversationSearchQueue.Enqueue(conversationId);
                            _searchDebounceTimer?.Stop();
                            _searchDebounceTimer?.Start();
                            continue;
                        }

                        var snapshot = BuildSnapshot(mailItem, out parentFolder);
                        if (snapshot is null)
                        {
                            DebugLogger.Log($"CatchUpConversationAsync snapshot null Event='{eventId}' Conversation='{conversationId}' Entry='{reference.EntryId}'");
                            continue;
                        }

                        await _eventRepository.TryAddMailAsync(snapshot, eventId).ConfigureAwait(false);
                        DebugLogger.Log($"CatchUpConversationAsync appended Entry='{snapshot.EntryId}' Event='{eventId}' Conversation='{conversationId}'");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Catch-up failed for Event='{eventId}' Conversation='{conversationId}' Entry='{reference.EntryId}': {ex.Message}");
                    }
                    finally
                    {
                        if (parentFolder is not null)
                        {
                            Marshal.ReleaseComObject(parentFolder);
                            parentFolder = null;
                        }

                        if (mailItem is not null)
                        {
                            Marshal.ReleaseComObject(mailItem);
                            mailItem = null;
                        }
                    }
                }
            }

            private IReadOnlyList<(string EntryId, string StoreId)> CollectConversationMailReferences(Outlook.NameSpace session, string eventId, string conversationId, DateTime cutoffUtc, CatchUpMetadata? metadata)
            {
                var matches = new List<(string EntryId, string StoreId)>();
                if (session is null || string.IsNullOrWhiteSpace(conversationId))
                {
                    return matches;
                }

                var seenEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var conversationStores = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (metadata?.ConversationEntryIds is not null)
                {
                    var normalizedConversation = conversationId.Trim();
                    if (metadata.ConversationEntryIds.TryGetValue(normalizedConversation, out var trackedEntries))
                    {
                        foreach (var tracked in trackedEntries)
                        {
                            if (!string.IsNullOrWhiteSpace(tracked))
                            {
                                seenEntries.Add(tracked.Trim());
                            }
                        }
                    }
                }

                var topicSet = metadata?.ConversationTopics;
                var messageIdSet = metadata?.MessageIds;
                var foundMessageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var conversationStatus = CollectConversationMatchesUsingConversationApi(session, eventId, conversationId, cutoffUtc, metadata, matches, seenEntries, conversationStores);
                
                var needsBroadSearch = (topicSet is not null && topicSet.Count > 0) || (messageIdSet is not null && messageIdSet.Count > 0);
                if (!needsBroadSearch && (conversationStatus.IsComplete || conversationStatus.Added > 0))
                {
                    return matches;
                }

                Outlook.Stores? stores = null;
                ISet<string>? normalizedFilter = null;
                var storeFilter = metadata?.StoreIds;

                if (storeFilter is null || storeFilter.Count == 0)
                {
                    Outlook.Store? defaultStore = null;
                    try
                    {
                        defaultStore = session.DefaultStore;
                    }
                    catch (COMException ex)
                    {
                        DebugLogger.Log($"CollectConversationMailReferences failed to access default store: {ex.Message}");
                        defaultStore = null;
                    }

                    if (defaultStore is null)
                    {
                        return matches;
                    }

                    try
                    {
                        var defaultStoreId = defaultStore.StoreID ?? string.Empty;
                        CollectMatchesFromStore(defaultStore, defaultStoreId, conversationId, cutoffUtc, matches, seenEntries, topicSet, messageIdSet, foundMessageIds);
                        
                        // Explicitly check Sent Items folder in default store
                        var sentFolder = SafeGetDefaultFolder(defaultStore, Outlook.OlDefaultFolders.olFolderSentMail);
                        if (sentFolder != null)
                        {
                            CollectMatchesFromFolder(sentFolder, defaultStoreId, conversationId, cutoffUtc, matches, seenEntries, false, topicSet, messageIdSet, foundMessageIds);
                            Marshal.ReleaseComObject(sentFolder);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(defaultStore);
                    }

                    CheckMissingMessageIds(eventId, messageIdSet, foundMessageIds);
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
                catch (Exception ex)
                {
                    DebugLogger.Log($"CollectConversationMailReferences store filter normalization failed: {ex.Message}");
                    normalizedFilter = null;
                }

                if (normalizedFilter is null)
                {
                    return matches;
                }

                if (conversationStores.Count > 0)
                {
                    foreach (var storeId in conversationStores)
                    {
                        if (!string.IsNullOrWhiteSpace(storeId))
                        {
                            normalizedFilter.Add(storeId.Trim());
                        }
                    }
                }

                Outlook.Store? fallbackDefaultStore = null;
                try
                {
                    fallbackDefaultStore = session.DefaultStore;
                    var defaultStoreId = fallbackDefaultStore?.StoreID;
                    var filter = normalizedFilter;
                    var storeId = defaultStoreId;
                    if (!string.IsNullOrWhiteSpace(storeId) && filter is not null)
                    {
                        filter.Add(storeId!.Trim());
                    }
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"CollectConversationMailReferences failed to access default store: {ex.Message}");
                }
                finally
                {
                    if (fallbackDefaultStore is not null)
                    {
                        Marshal.ReleaseComObject(fallbackDefaultStore);
                    }
                }

                if (normalizedFilter.Count == 0)
                {
                    return matches;
                }

                try
                {
                    stores = session.Stores;
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"CollectConversationMailReferences failed to access stores: {ex.Message}");
                    return matches;
                }

                if (stores is null)
                {
                    return matches;
                }

                var remaining = new HashSet<string>(normalizedFilter, StringComparer.OrdinalIgnoreCase);

                try
                {
                    var count = stores.Count;
                    for (var index = 1; index <= count && remaining.Count > 0; index++)
                    {
                        Outlook.Store? store = null;
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
                            CollectMatchesFromStore(store, storeId, conversationId, cutoffUtc, matches, seenEntries, topicSet, messageIdSet, foundMessageIds);
                        }
                        finally
                        {
                            if (store is not null)
                            {
                                Marshal.ReleaseComObject(store);
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(stores);
                }

                CheckMissingMessageIds(eventId, messageIdSet, foundMessageIds);
                return matches;
            }

            private void CheckMissingMessageIds(string eventId, ISet<string>? requested, ISet<string>? found)
            {
                if (requested is null || requested.Count == 0)
                {
                    return;
                }

                var missing = new HashSet<string>(requested, StringComparer.OrdinalIgnoreCase);
                if (found is not null)
                {
                    missing.ExceptWith(found);
                }

                if (missing.Count > 0)
                {
                    DebugLogger.Log($"CollectConversationMailReferences marking {missing.Count} message IDs as not found for Event='{eventId}'");
                    _ = _eventRepository.MarkMessageIdsAsNotFoundAsync(eventId, missing);
                }
            }

            private ConversationCaptureStatus CollectConversationMatchesUsingConversationApi(Outlook.NameSpace session, string eventId, string conversationId, DateTime cutoffUtc, CatchUpMetadata? metadata, IList<(string EntryId, string StoreId)> matches, ISet<string> seen, ISet<string> capturedStores)
            {
                try
                {
                    if (session is null || metadata?.ConversationEntryIds is null || metadata.ConversationEntryIds.Count == 0)
                    {
                        return default;
                    }

                    var normalizedConversationId = (conversationId ?? string.Empty).Trim();
                    if (normalizedConversationId.Length == 0)
                    {
                        return default;
                    }

                    if (!metadata.ConversationEntryIds.TryGetValue(normalizedConversationId, out var candidateEntryIds) || candidateEntryIds.Count == 0)
                    {
                        return default;
                    }

                    foreach (var candidateEntryId in candidateEntryIds)
                    {
                        if (string.IsNullOrWhiteSpace(candidateEntryId))
                        {
                            continue;
                        }

                        Outlook.MailItem? root = null;
                        Outlook.Conversation? conversation = null;

                        try
                        {
                            root = ResolveMailItem(session, candidateEntryId);
                            if (root is null)
                            {
                                continue;
                            }

                            conversation = root.GetConversation();
                            if (conversation is null)
                            {
                                continue;
                            }

                            var enumerationResult = EnumerateConversationEntries(conversation, normalizedConversationId, cutoffUtc);
                            if (enumerationResult.TotalCount == 0)
                            {
                                continue;
                            }

                            var added = 0;
                            var alreadyTracked = 0;

                            foreach (var entry in enumerationResult.Entries)
                            {
                                if (string.IsNullOrWhiteSpace(entry.EntryId))
                                {
                                    continue;
                                }

                                var storeIdCandidate = entry.StoreId;
                                if (!string.IsNullOrWhiteSpace(storeIdCandidate))
                                {
                                    capturedStores.Add(storeIdCandidate!);
                                }

                                if (!seen.Add(entry.EntryId))
                                {
                                    alreadyTracked++;
                                    continue;
                                }

                                matches.Add((entry.EntryId, entry.StoreId ?? string.Empty));
                                added++;
                            }

                            var coverage = added + alreadyTracked;
                            var isComplete = enumerationResult.TotalCount > 0 && coverage >= enumerationResult.TotalCount;

                            ScheduleConversationCompletenessEvaluation(eventId, normalizedConversationId, enumerationResult, coverage);

                            if (isComplete)
                            {
                                DebugLogger.Log($"CollectConversationMailReferences conversation completeness satisfied Event='{eventId}' Conversation='{normalizedConversationId}' Coverage={coverage}/{enumerationResult.TotalCount}");
                            }
                            else if (added > 0)
                            {
                                DebugLogger.Log($"CollectConversationMailReferences captured {coverage}/{enumerationResult.TotalCount} conversation items Event='{eventId}' Conversation='{normalizedConversationId}' Added={added}");
                            }

                            if (added > 0 || isComplete)
                            {
                                return new ConversationCaptureStatus(added, coverage, enumerationResult.TotalCount, isComplete);
                            }
                        }
                        catch (COMException ex)
                        {
                            DebugLogger.Log($"CollectConversationMailReferences conversation API failed for seed entry '{candidateEntryId}': {ex.Message}");
                        }
                        finally
                        {
                            if (conversation is not null)
                            {
                                Marshal.ReleaseComObject(conversation);
                                conversation = null;
                            }

                            if (root is not null)
                            {
                                Marshal.ReleaseComObject(root);
                                root = null;
                            }
                        }
                    }

                    return default;
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"CollectConversationMatchesUsingConversationApi failed Event='{eventId}' Conversation='{conversationId}': {ex}");
                    return default;
                }
            }

            private ConversationEnumerationResult EnumerateConversationEntries(Outlook.Conversation conversation, string normalizedConversationId, DateTime cutoffUtc)
            {
                if (conversation is null)
                {
                    return ConversationEnumerationResult.Empty;
                }

                Outlook.Table? table = null;

                try
                {
                    table = conversation.GetTable();

                    TryAddConversationColumn(table, "EntryID");
                    TryAddConversationColumn(table, "StoreID");
                    TryAddConversationColumn(table, "ConversationID");
                    TryAddConversationColumn(table, "ReceivedTime");

                    var entries = new List<ConversationEntry>();

                    while (!table.EndOfTable)
                    {
                        Outlook.Row row;
                        try
                        {
                            row = table.GetNextRow();
                        }
                        catch (COMException ex)
                        {
                            DebugLogger.Log($"EnumerateConversationEntries failed to advance table: {ex.Message}");
                            break;
                        }

                        if (row is null)
                        {
                            break;
                        }

                        var entryId = GetRowString(row, "EntryID");
                        if (string.IsNullOrWhiteSpace(entryId))
                        {
                            continue;
                        }

                        var normalizedEntryId = entryId!.Trim();
                        if (normalizedEntryId.Length == 0)
                        {
                            continue;
                        }

                        var conversationId = GetRowString(row, "ConversationID");
                        if (!string.IsNullOrWhiteSpace(conversationId) &&
                            !string.Equals(conversationId, normalizedConversationId, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var receivedUtc = GetRowUtc(row, "ReceivedTime");
                        if (receivedUtc.HasValue && receivedUtc.Value < cutoffUtc)
                        {
                            continue;
                        }

                        var storeId = GetRowString(row, "StoreID");
                        entries.Add(new ConversationEntry(normalizedEntryId, storeId?.Trim(), receivedUtc));
                    }

                    return entries.Count == 0
                        ? ConversationEnumerationResult.Empty
                        : new ConversationEnumerationResult(entries);
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"EnumerateConversationEntries failed: {ex.Message}");
                    return ConversationEnumerationResult.Empty;
                }
                finally
                {
                    if (table is not null)
                    {
                        Marshal.ReleaseComObject(table);
                    }
                }
            }

            private void ScheduleConversationCompletenessEvaluation(string eventId, string conversationId, ConversationEnumerationResult enumerationResult, int coverage)
            {
                if (string.IsNullOrWhiteSpace(eventId) || enumerationResult.TotalCount == 0)
                {
                    return;
                }

                _ = Task.Run(async () =>
                {
                    try
                    {
                        await EvaluateConversationCompletenessAsync(eventId, conversationId, enumerationResult, coverage).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"ScheduleConversationCompletenessEvaluation failed Event='{eventId}' Conversation='{conversationId}': {ex.Message}");
                    }
                });
            }

            private async Task EvaluateConversationCompletenessAsync(string eventId, string conversationId, ConversationEnumerationResult enumerationResult, int coverage)
            {
                try
                {
                    var record = await _eventRepository.GetByIdAsync(eventId).ConfigureAwait(false);
                    if (record is null)
                    {
                        return;
                    }

                    var normalizedConversationId = (conversationId ?? string.Empty).Trim();

                    var tracked = record.Emails
                        .Where(email => string.Equals(email.ConversationId, normalizedConversationId, StringComparison.OrdinalIgnoreCase))
                        .Select(email => email.EntryId)
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    var missing = enumerationResult.Entries
                        .Select(entry => entry.EntryId)
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Where(id => !tracked.Contains(id))
                        .ToList();

                    if (missing.Count == 0)
                    {
                        DebugLogger.Log($"Conversation completeness check satisfied Event='{eventId}' Conversation='{normalizedConversationId}' Coverage={coverage}/{enumerationResult.TotalCount}");
                        return;
                    }

                    var preview = string.Join(",", missing.Take(10));
                    if (missing.Count > 10)
                    {
                        preview += ",...";
                    }

                    DebugLogger.Log($"Conversation completeness check pending Event='{eventId}' Conversation='{normalizedConversationId}' MissingCount={missing.Count} PendingEntryIds={preview}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Conversation completeness evaluation failed Event='{eventId}' Conversation='{conversationId}': {ex}");
                }
            }

            private static void TryAddConversationColumn(Outlook.Table? table, string columnName)
            {
                if (table is null || string.IsNullOrWhiteSpace(columnName))
                {
                    return;
                }

                try
                {
                    table.Columns.Add(columnName);
                }
                catch (COMException)
                {
                    // column might already exist; ignore
                }
                catch (ArgumentException)
                {
                    // column might already exist; ignore
                }
                catch (InvalidCastException)
                {
                    // some stores throw when adding built-in columns; ignore
                }
            }

            private static string? GetRowString(Outlook.Row row, string columnName)
            {
                if (row is null || string.IsNullOrWhiteSpace(columnName))
                {
                    return null;
                }

                try
                {
                    var value = row[columnName];
                    return value switch
                    {
                        null => null,
                        string text => text,
                        _ => value.ToString()
                    };
                }
                catch (COMException)
                {
                    return null;
                }
                catch (ArgumentException)
                {
                    return null;
                }
                catch (InvalidCastException)
                {
                    return null;
                }
            }

            private static DateTime? GetRowUtc(Outlook.Row row, string columnName)
            {
                if (row is null || string.IsNullOrWhiteSpace(columnName))
                {
                    return null;
                }

                try
                {
                    var value = row[columnName];
                    if (value is DateTime dt)
                    {
                        return ToUniversal(dt);
                    }

                    if (value is string text && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
                    {
                        return ToUniversal(parsed);
                    }

                    if (value is double oaDate)
                    {
                        return ToUniversal(DateTime.FromOADate(oaDate));
                    }
                }
                catch (COMException)
                {
                    return null;
                }
                catch (ArgumentException)
                {
                    return null;
                }
                catch (InvalidCastException)
                {
                    return null;
                }

                return null;
            }

            private void CollectMatchesFromStore(Outlook.Store store, string storeId, string conversationId, DateTime cutoffUtc, IList<(string EntryId, string StoreId)> matches, ISet<string> seen, ISet<string>? topicSet, ISet<string>? messageIdSet, ISet<string>? foundMessageIds)
            {
                if (store is null)
                {
                    return;
                }

                var inbox = SafeGetDefaultFolder(store, Outlook.OlDefaultFolders.olFolderInbox);
                var sent = SafeGetDefaultFolder(store, Outlook.OlDefaultFolders.olFolderSentMail);
                var deleted = SafeGetDefaultFolder(store, Outlook.OlDefaultFolders.olFolderDeletedItems);

                try
                {
                    CollectMatchesFromFolder(inbox, storeId, conversationId, cutoffUtc, matches, seen, true, topicSet, messageIdSet, foundMessageIds);
                    CollectMatchesFromFolder(sent, storeId, conversationId, cutoffUtc, matches, seen, false, topicSet, messageIdSet, foundMessageIds);
                    CollectMatchesFromFolder(deleted, storeId, conversationId, cutoffUtc, matches, seen, false, topicSet, messageIdSet, foundMessageIds);
                }
                finally
                {
                    if (deleted is not null)
                    {
                        Marshal.ReleaseComObject(deleted);
                    }

                    if (sent is not null)
                    {
                        Marshal.ReleaseComObject(sent);
                    }

                    if (inbox is not null)
                    {
                        Marshal.ReleaseComObject(inbox);
                    }
                }
            }

            private async Task<CatchUpMetadata?> GetCatchUpMetadataAsync(string eventId)
            {
                if (string.IsNullOrWhiteSpace(eventId))
                {
                    return null;
                }

                try
                {
                    var record = await _eventRepository.GetByIdAsync(eventId);
                    if (record is null)
                    {
                        return null;
                    }

                    var allEmails = record.Emails ?? new List<EmailItem>();

                    var stores = allEmails
                        .Where(email => !string.IsNullOrWhiteSpace(email.StoreId))
                        .Select(email => email.StoreId.Trim())
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var topicSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var messageIdSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var existingMessageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var conversationEntryMap = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                    DateTime? earliestReceivedUtc = null;

                    void AddTopic(string? subject)
                    {
                        var normalized = NormalizeConversationTopic(subject);
                        if (!string.IsNullOrEmpty(normalized))
                        {
                            topicSet.Add(normalized);
                        }
                    }

                    AddTopic(record.EventTitle);

                    // Add explicitly tracked related subjects from the event record
                    if (record.RelatedSubjects != null)
                    {
                        foreach (var subject in record.RelatedSubjects)
                        {
                            AddTopic(subject);
                        }
                    }

                    foreach (var email in allEmails)
                    {
                        AddTopic(email.Subject);

                        if (!string.IsNullOrWhiteSpace(email.InternetMessageId))
                        {
                            foreach (var id in ExtractMessageIds(email.InternetMessageId))
                            {
                                messageIdSet.Add(id);
                                existingMessageIds.Add(id);
                            }
                        }

                        if (email.ReferenceMessageIds is not null && email.ReferenceMessageIds.Length > 0)
                        {
                            foreach (var referenceId in email.ReferenceMessageIds)
                            {
                                foreach (var id in ExtractMessageIds(referenceId))
                                {
                                    messageIdSet.Add(id);
                                }
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(email.ConversationId) && !string.IsNullOrWhiteSpace(email.EntryId))
                        {
                            var conversationKey = email.ConversationId.Trim();
                            var entryKey = email.EntryId.Trim();
                            if (conversationKey.Length > 0 && entryKey.Length > 0)
                            {
                                if (!conversationEntryMap.TryGetValue(conversationKey, out var entrySet))
                                {
                                    entrySet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                    conversationEntryMap[conversationKey] = entrySet;
                                }

                                entrySet.Add(entryKey);
                            }
                        }

                        var receivedUtc = ToUniversal(email.ReceivedOn);
                        if (receivedUtc.HasValue && (!earliestReceivedUtc.HasValue || receivedUtc.Value < earliestReceivedUtc.Value))
                        {
                            earliestReceivedUtc = receivedUtc;
                        }
                    }

                    // Do NOT filter out existing emails from the search set.
                    // We need them in messageIdSet so that NEW emails (replies) that reference existing emails
                    // can be identified and linked.
                    // messageIdSet.ExceptWith(existingMessageIds);

                    // Filter out known missing emails
                    if (record.NotFoundMessageIds != null)
                    {
                        messageIdSet.ExceptWith(record.NotFoundMessageIds);
                    }

                    IReadOnlyDictionary<string, IReadOnlyCollection<string>>? conversationEntries = null;
                    if (conversationEntryMap.Count > 0)
                    {
                        var map = new Dictionary<string, IReadOnlyCollection<string>>(conversationEntryMap.Count, StringComparer.OrdinalIgnoreCase);
                        foreach (var kvp in conversationEntryMap)
                        {
                            map[kvp.Key] = kvp.Value.ToList();
                        }

                        conversationEntries = map;
                    }

                    return new CatchUpMetadata
                    {
                        StoreIds = stores.Count > 0 ? stores : Array.Empty<string>(),
                        ConversationTopics = topicSet.Count > 0 ? topicSet : null,
                        MessageIds = messageIdSet.Count > 0 ? messageIdSet : null,
                        NotFoundMessageIds = record.NotFoundMessageIds,
                        EarliestReceivedOnUtc = earliestReceivedUtc,
                        ConversationEntryIds = conversationEntries
                    };
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"GetCatchUpMetadataAsync failed for Event='{eventId}': {ex.Message}");
                    return null;
                }
            }

            private void CollectMatchesFromFolder(Outlook.MAPIFolder? folder, string storeId, string conversationId, DateTime cutoffUtc, IList<(string EntryId, string StoreId)> matches, ISet<string> seen, bool includeChildren, ISet<string>? topicSet, ISet<string>? messageIdSet, ISet<string>? foundMessageIds)
            {
                if (folder is null)
                {
                    return;
                }

                try
                {
                    if (folder.DefaultItemType != Outlook.OlItemType.olMailItem)
                    {
                        return;
                    }
                }
                catch (COMException)
                {
                    return;
                }

                Outlook.Items? items = null;
                try
                {
                    items = folder.Items;
                    if (items is null)
                    {
                        return;
                    }

                    items.Sort("[ReceivedTime]", true);

                    bool processedConversationFilter = false;
                    var needsBroadSearch = (topicSet is not null && topicSet.Count > 0)
                                            || (messageIdSet is not null && messageIdSet.Count > 0);

                    void ProcessRestrictedItems(Outlook.Items? view)
                    {
                        if (view is null)
                        {
                            return;
                        }

                        object? currentItem = null;

                        try
                        {
                            currentItem = view.GetFirst();
                            while (currentItem is not null)
                            {
                                if (currentItem is Outlook.MailItem mail)
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

                                        var conversationMatch = !string.IsNullOrEmpty(mailConversationId) &&
                                                                string.Equals(mailConversationId, conversationId, StringComparison.OrdinalIgnoreCase);

                                        var mailTopicNormalized = string.Empty;
                                        var matchesTopic = false;
                                        if (topicSet is not null && topicSet.Count > 0)
                                        {
                                            try
                                            {
                                                var topic = mail.ConversationTopic ?? string.Empty;
                                                mailTopicNormalized = NormalizeConversationTopic(topic);
                                                matchesTopic = !string.IsNullOrEmpty(mailTopicNormalized) && topicSet.Any(t => IsTopicMatch(t, mailTopicNormalized));
                                            }
                                            catch (COMException)
                                            {
                                                matchesTopic = false;
                                            }
                                        }

                                        var matchesReference = false;
                                        var matchedReferenceId = string.Empty;
                                        var matchedByMessageId = false;
                                        var candidateMessageId = GetInternetMessageId(mail);
                                        
                                        // Temporary Debugging for TEST_FAIL
                                        var debugSubject = mail.Subject;
                                        var isDebugTarget = !string.IsNullOrEmpty(debugSubject) && debugSubject.Contains("TEST_FAIL");

                                        if (!string.IsNullOrEmpty(candidateMessageId) && messageIdSet is not null && messageIdSet.Count > 0)
                                        {
                                            foreach (var id in ExtractMessageIds(candidateMessageId))
                                            {
                                                if (messageIdSet.Contains(id))
                                                {
                                                    matchesReference = true;
                                                    matchedReferenceId = id;
                                                    matchedByMessageId = true;
                                                    foundMessageIds?.Add(id);
                                                    break;
                                                }
                                            }
                                        }

                                        if (!matchesReference && messageIdSet is not null && messageIdSet.Count > 0)
                                        {
                                            var referenceIds = CollectReferenceMessageIds(mail);
                                            
                                            if (isDebugTarget)
                                            {
                                                var refString = string.Join(", ", referenceIds);
                                                DebugLogger.Log($"[DEBUG] Checking mail '{debugSubject}' Entry='{entryId}'. CandidateID='{candidateMessageId}'. Refs='{refString}'. MessageIdSetCount={messageIdSet.Count}");
                                                DebugLogger.Log($"[DEBUG] MessageIdSet: {string.Join(", ", messageIdSet)}");
                                            }

                                            foreach (var referenceId in referenceIds)
                                            {
                                                if (messageIdSet.Contains(referenceId))
                                                {
                                                    matchesReference = true;
                                                    matchedReferenceId = referenceId;
                                                    foundMessageIds?.Add(referenceId);
                                                    break;
                                                }
                                            }
                                        }

                                        if (!conversationMatch && !matchesReference && !matchesTopic && topicSet is not null && topicSet.Count > 0)
                                        {
                                            try
                                            {
                                                if (string.IsNullOrEmpty(mailTopicNormalized))
                                                {
                                                    var topic = mail.Subject ?? string.Empty;
                                                    mailTopicNormalized = NormalizeConversationTopic(topic);
                                                }

                                                matchesTopic = !string.IsNullOrEmpty(mailTopicNormalized) && topicSet.Any(t => IsTopicMatch(t, mailTopicNormalized));
                                            }
                                            catch (COMException)
                                            {
                                                matchesTopic = false;
                                            }
                                        }

                                        var shouldInclude = conversationMatch || matchesReference || matchesTopic;

                                        if (!shouldInclude || string.IsNullOrEmpty(entryId))
                                        {
                                            continue;
                                        }

                                        if (seen.Add(entryId))
                                        {
                                            if (matchesTopic && !conversationMatch)
                                            {
                                                DebugLogger.Log($"CollectMatchesFromFolder matched by topic Entry='{entryId}' Topic='{mailTopicNormalized}' Conversation='{conversationId}' Folder='{folder.Name}'");
                                            }

                                            if (matchesReference && !conversationMatch)
                                            {
                                                var referenceType = matchedByMessageId ? "message id" : "reference";
                                                DebugLogger.Log($"CollectMatchesFromFolder matched by {referenceType} Entry='{entryId}' Reference='{matchedReferenceId}' Conversation='{conversationId}' Folder='{folder.Name}'");
                                            }

                                            matches.Add((entryId, storeId));
                                        }
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(mail);
                                        currentItem = null;
                                    }
                                }
                                else
                                {
                                    Marshal.ReleaseComObject(currentItem);
                                    currentItem = null;
                                }

                                currentItem = view.GetNext();
                            }
                        }
                        finally
                        {
                            if (currentItem is not null)
                            {
                                Marshal.ReleaseComObject(currentItem);
                            }
                        }
                    }

                    if (s_conversationFilterSupported)
                    {
                        Outlook.Items? conversationFiltered = null;
                        try
                        {
                            var conversationFilter = BuildConversationFilter(conversationId, cutoffUtc);
                            conversationFiltered = items.Restrict(conversationFilter);
                        }
                        catch (COMException ex)
                        {
                            s_conversationFilterSupported = false;
                            DebugLogger.Log($"CollectMatchesFromFolder conversation filter unsupported Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}. Falling back to received-time filter.");
                        }

                        if (conversationFiltered is not null)
                        {
                            processedConversationFilter = true;
                            try
                            {
                                ProcessRestrictedItems(conversationFiltered);
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(conversationFiltered);
                            }
                        }
                    }

                    // New Logic: Subject-based filtering (User requested to prioritize Subject over MessageID)
                    if (topicSet is not null && topicSet.Count > 0)
                    {
                        var subjectFilter = BuildSubjectFilter(topicSet);
                        if (!string.IsNullOrEmpty(subjectFilter))
                        {
                            Outlook.Items? subjectView = null;
                            try
                            {
                                subjectView = items.Restrict(subjectFilter);
                                if (subjectView is not null)
                                {
                                    DebugLogger.Log($"CollectMatchesFromFolder applying subject filter Folder='{folder.Name}' Conversation='{conversationId}' Topics={topicSet.Count}");
                                    ProcessRestrictedItems(subjectView);
                                }
                            }
                            catch (COMException ex)
                            {
                                DebugLogger.Log($"CollectMatchesFromFolder subject filter failed Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}");
                            }
                            finally
                            {
                                if (subjectView is not null)
                                {
                                    Marshal.ReleaseComObject(subjectView);
                                }
                            }
                        }
                    }

                    // Disabled MessageID-based filtering as per user request ("Search using Subject, not MessageID")
                    /*
                    if (messageIdSet is not null && messageIdSet.Count > 0)
                    {
                        var referenceFilter = BuildReferenceFilter(messageIdSet);
                        if (!string.IsNullOrEmpty(referenceFilter))
                        {
                            Outlook.Items? referenceView = null;
                            try
                            {
                                referenceView = items.Restrict(referenceFilter);
                                if (referenceView is not null)
                                {
                                    var appliedCount = Math.Min(messageIdSet.Count, MaxReferenceFilterTerms);
                                    DebugLogger.Log($"CollectMatchesFromFolder applying reference filter Folder='{folder.Name}' Conversation='{conversationId}' MessageIds={appliedCount}");
                                    ProcessRestrictedItems(referenceView);
                                }
                            }
                            catch (COMException ex)
                            {
                                DebugLogger.Log($"CollectMatchesFromFolder reference filter failed Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}");
                            }
                            finally
                            {
                                if (referenceView is not null)
                                {
                                    Marshal.ReleaseComObject(referenceView);
                                }
                            }
                        }
                    }
                    */

                    if (!processedConversationFilter || !s_conversationFilterSupported || needsBroadSearch)
                    {
                        Outlook.Items? fallbackView = null;
                        try
                        {
                            var fallbackFilter = BuildReceivedSinceFilter(cutoffUtc);
                            fallbackView = items.Restrict(fallbackFilter);
                            if (fallbackView is not null)
                            {
                                if (processedConversationFilter && needsBroadSearch)
                                {
                                    DebugLogger.Log($"CollectMatchesFromFolder expanding search by received-time filter Folder='{folder.Name}' Conversation='{conversationId}'");
                                }

                                ProcessRestrictedItems(fallbackView);
                            }
                        }
                        catch (COMException ex)
                        {
                            DebugLogger.Log($"CollectMatchesFromFolder received-time filter failed Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}");
                        }
                        finally
                        {
                            if (fallbackView is not null)
                            {
                                Marshal.ReleaseComObject(fallbackView);
                            }
                        }
                    }
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"CollectMatchesFromFolder failed Folder='{folder.Name}' Conversation='{conversationId}': {ex.Message}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"CollectMatchesFromFolder unexpected exception Folder='{folder.Name}' Conversation='{conversationId}': {ex}");
                }
                finally
                {
                    if (items is not null)
                    {
                        Marshal.ReleaseComObject(items);
                    }
                }

                if (!includeChildren)
                {
                    return;
                }

                Outlook.Folders? children = null;
                try
                {
                    children = folder.Folders;
                    if (children is null || children.Count == 0)
                    {
                        return;
                    }

                    for (var index = 1; index <= children.Count; index++)
                    {
                        Outlook.MAPIFolder? child = null;
                        try
                        {
                            child = children[index];
                            if (child is null)
                            {
                                continue;
                            }

                            CollectMatchesFromFolder(child, storeId, conversationId, cutoffUtc, matches, seen, true, topicSet, messageIdSet, foundMessageIds);
                        }
                        finally
                        {
                            if (child is not null)
                            {
                                Marshal.ReleaseComObject(child);
                            }
                        }
                    }
                }
                finally
                {
                    if (children is not null)
                    {
                        Marshal.ReleaseComObject(children);
                    }
                }
            }

            private static Outlook.MAPIFolder? SafeGetDefaultFolder(Outlook.Store store, Outlook.OlDefaultFolders folderType)
            {
                if (store is null)
                {
                    return null;
                }

                try
                {
                    return store.GetDefaultFolder(folderType);
                }
                catch (COMException)
                {
                    return null;
                }
            }

            private static string BuildConversationFilter(string conversationId, DateTime cutoffUtc)
            {
                var cutoffLocal = cutoffUtc.ToLocalTime();
                var culture = CultureInfo.CreateSpecificCulture("en-US");
                var escapedConversationId = conversationId.Replace("'", "''");
                return $"[ConversationID] = '{escapedConversationId}' AND [ReceivedTime] >= \"{cutoffLocal.ToString("g", culture)}\"";
            }

            private const int MaxReferenceFilterTerms = 10;

            private static string BuildReceivedSinceFilter(DateTime cutoffUtc)
            {
                var cutoffLocal = cutoffUtc.ToLocalTime();
                var culture = CultureInfo.CreateSpecificCulture("en-US");
                // Include Sent Items by checking CreationTime as well, since Sent Items might not have ReceivedTime set correctly or at all in some contexts?
                // Actually, Sent Items have SentOn or CreationTime. ReceivedTime usually maps to SentOn for sent items.
                // But to be safe, let's just use ReceivedTime as it is the standard property for "date" in filters.
                return $"[ReceivedTime] >= \"{cutoffLocal.ToString("g", culture)}\"";
            }

            private static string? BuildReferenceFilter(ISet<string>? messageIdSet)
            {
                if (messageIdSet is null || messageIdSet.Count == 0)
                {
                    return null;
                }

                var ids = messageIdSet
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Take(MaxReferenceFilterTerms)
                    .Select(id => id.Trim())
                    .ToList();

                if (ids.Count == 0)
                {
                    return null;
                }

                var clauses = new List<string>(ids.Count);
                foreach (var id in ids)
                {
                    var escaped = id.Replace("'", "''");
                    clauses.Add($"(\"urn:schemas:mailheader:references\" LIKE '%{escaped}%' OR \"urn:schemas:mailheader:in-reply-to\" LIKE '%{escaped}%' OR \"http://schemas.microsoft.com/mapi/proptag/0x1035001F\" = '{escaped}')");
                }

                return clauses.Count == 0 ? null : "@SQL=" + string.Join(" OR ", clauses);
            }

            private static string? BuildSubjectFilter(ISet<string>? topicSet)
            {
                if (topicSet is null || topicSet.Count == 0)
                {
                    return null;
                }

                var topics = topicSet
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Where(t => t.Length < 255)
                    .Take(MaxReferenceFilterTerms) // Reuse limit or define new one
                    .Select(t => t.Trim())
                    .ToList();

                if (topics.Count == 0)
                {
                    return null;
                }

                var clauses = new List<string>(topics.Count);
                foreach (var topic in topics)
                {
                    // Replace non-alphanumeric characters with space
                    var cleanTopic = Regex.Replace(topic, @"[^\p{L}\p{N}]", " ");

                    // Split into words to avoid phrase matching issues with punctuation (e.g. "PO." vs "PO")
                    // Using AND logic ensures all words are present without enforcing adjacency/punctuation strictness
                    var words = cleanTopic.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (words.Length == 0)
                    {
                        continue;
                    }

                    // Strategy 1: Token-based AND match (for English/Space-separated languages)
                    // Limit number of words to avoid query overflow and handle subject truncation
                    // Reduced from 10 to 5 to ensure we don't require words that might be truncated
                    var selectedWords = words.Take(5).ToList();
                    var wordClauses = new List<string>();
                    
                    // If we have very few words (e.g. Chinese title treated as 1 word), allow prefix matching on the last word
                    // to handle truncation (e.g. "关于项目..." vs "关于项目A的讨论")
                    bool allowPrefix = words.Length <= 3;

                    for (int i = 0; i < selectedWords.Count; i++)
                    {
                        var word = selectedWords[i];
                        var escapedWord = word.Replace("'", "''");
                        
                        // Apply prefix wildcard only to the last word if we have few words
                        if (allowPrefix && i == selectedWords.Count - 1)
                        {
                             wordClauses.Add($"\"urn:schemas:httpmail:subject\" ci_phrasematch '{escapedWord}*'");
                        }
                        else
                        {
                             wordClauses.Add($"\"urn:schemas:httpmail:subject\" ci_phrasematch '{escapedWord}'");
                        }
                    }

                    if (wordClauses.Count > 0)
                    {
                        clauses.Add($"({string.Join(" AND ", wordClauses)})");
                    }
                }

                return "@SQL=" + string.Join(" OR ", clauses);
            }

            private static readonly string[] SubjectPrefixes = new[] { "RE:", "FW:", "FWD:", "回复:", "转发:", "回覆:", "轉寄:" };

            private static string NormalizeConversationTopic(string? topic)
            {
                if (string.IsNullOrWhiteSpace(topic))
                {
                    return string.Empty;
                }

                var normalized = (topic ?? string.Empty).Trim();
                bool changed = true;
                while (changed)
                {
                    changed = false;
                    foreach (var prefix in SubjectPrefixes)
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

            private static bool IsTopicMatch(string storedTopic, string mailTopic)
            {
                if (string.IsNullOrEmpty(storedTopic) || string.IsNullOrEmpty(mailTopic))
                    return false;

                // Exact match (normalized)
                if (string.Equals(storedTopic, mailTopic, StringComparison.OrdinalIgnoreCase))
                    return true;

                // Handle truncation: storedTopic starts with mailTopic
                // Example: storedTopic="Project Alpha Final Report", mailTopic="Project Alpha Fin"
                // We require mailTopic to be a prefix of storedTopic.
                // Also require mailTopic to have some substance to avoid matching "Re:" or "P"
                // 4 chars is reasonable for "Plan", "Test", "关于项目"
                if (mailTopic.Length >= 4 && storedTopic.StartsWith(mailTopic, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                
                // Handle reverse truncation (rare, but possible if stored topic was truncated and new email has full topic)
                // Example: storedTopic="Project Alpha", mailTopic="Project Alpha Final"
                if (storedTopic.Length >= 4 && mailTopic.StartsWith(storedTopic, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                return false;
                       }

            private async Task<Outlook.MailItem?> ResolveMailItemWithRetryAsync(Outlook.NameSpace session, string entryId, int maxAttempts = 5)
            {
                Outlook.MailItem? item = null;

                for (var attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    item = ResolveMailItem(session, entryId);
                    if (item is not null)
                    {
                        if (attempt > 1)
                        {
                            DebugLogger.Log($"ResolveMailItemWithRetryAsync succeeded for entry '{entryId}' on attempt {attempt}");
                        }

                        return item;
                    }

                    if (attempt < maxAttempts)
                    {
                        var delayMs = attempt switch
                        {
                            1 => 400,
                            2 => 800,
                            3 => 1600,
                            4 => 3200,
                            _ => 0
                        };

                        if (delayMs > 0)
                        {
                            DebugLogger.Log($"ResolveMailItemWithRetryAsync retrying entry '{entryId}' in {delayMs}ms (attempt {attempt})");
                            await Task.Delay(delayMs).ConfigureAwait(false);
                        }
                    }
                }

                DebugLogger.Log($"ResolveMailItemWithRetryAsync failed to resolve entry '{entryId}' after {maxAttempts} attempts");
                return item;
            }

            private Outlook.NameSpace? EnsureSession()
            {
                if (_session is not null)
                {
                    return _session;
                }

                try
                {
                    _session = _application.Session;
                }
                catch (COMException ex)
                {
                    Debug.WriteLine($"[OutlookEventMonitor] Unable to get Outlook session: {ex}");
                    _session = null;
                }

                return _session;
            }

            private void DrainCatchUpQueue()
            {
                while (_catchUpQueue.TryDequeue(out _))
                {
                    // intentionally empty
                }
            }

            private MailSnapshot? BuildSnapshot(Outlook.MailItem mailItem, out Outlook.MAPIFolder? parentFolder)
            {
                parentFolder = null;

                if (mailItem is null)
                {
                    return null;
                }

                try
                {
                    parentFolder = mailItem.Parent as Outlook.MAPIFolder;
                }
                catch (COMException)
                {
                    parentFolder = null;
                }

                var conversationId = mailItem.ConversationID;
                if (string.IsNullOrWhiteSpace(conversationId))
                {
                    DebugLogger.Log($"BuildSnapshot missing conversation for entry '{mailItem.EntryID}'");
                    return null;
                }

                var threadIndex = GetThreadIndex(mailItem);
                var threadIndexPrefix = GetThreadIndexPrefix(mailItem, threadIndex);

                return new MailSnapshot
                {
                    EntryId = mailItem.EntryID ?? string.Empty,
                    StoreId = parentFolder?.StoreID ?? string.Empty,
                    ConversationId = conversationId,
                    InternetMessageId = GetInternetMessageId(mailItem),
                    Sender = mailItem.SenderName ?? string.Empty,
                    To = mailItem.To ?? string.Empty,
                    Subject = mailItem.Subject ?? string.Empty,
                    Participants = MailParticipantExtractor.Capture(mailItem),
                    BodyFingerprint = MailBodyFingerprint.Capture(mailItem),
                    ThreadIndex = threadIndex,
                    ThreadIndexPrefix = threadIndexPrefix,
                    ReferenceMessageIds = CollectReferenceMessageIds(mailItem),
                    HistoricalSubjects = ExtractHistoricalSubjects(mailItem),
                    ReceivedOn = mailItem.ReceivedTime.ToUniversalTime(),
                    Attachments = CaptureAttachments(mailItem)
                };
            }

            private static IReadOnlyList<string> ExtractHistoricalSubjects(Outlook.MailItem mailItem)
            {
                var subjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var body = mailItem.Body;
                if (string.IsNullOrWhiteSpace(body))
                {
                    return Array.Empty<string>();
                }

                // 1. Try standard extraction
                var matches = HistoricalSubjectRegex.Matches(body);
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        var rawSubject = match.Groups["subject"].Value;
                        var subject = Regex.Replace(rawSubject, @"\r?\n\s+", " ").Trim();
                        if (!string.IsNullOrWhiteSpace(subject))
                        {
                            subjects.Add(subject);
                        }
                    }
                }
                else
                {
                    // 2. Fallback: Try to repair the body if no matches found (Mojibake detection)
                    if (EncodingRepair.TryFix(body, s => HistoricalSubjectRegex.IsMatch(s), out var repairedBody))
                    {
                        foreach (Match match in HistoricalSubjectRegex.Matches(repairedBody))
                        {
                            var rawSubject = match.Groups["subject"].Value;
                            var subject = Regex.Replace(rawSubject, @"\r?\n\s+", " ").Trim();
                            if (!string.IsNullOrWhiteSpace(subject))
                            {
                                subjects.Add(subject);
                            }
                        }
                    }
                }

                return subjects.ToList();
            }

            private static IReadOnlyList<AttachmentItem> CaptureAttachments(Outlook.MailItem mailItem)
            {
                var attachments = new List<AttachmentItem>();

                if (mailItem.Attachments is null || mailItem.Attachments.Count == 0)
                {
                    return attachments;
                }

                var entryId = mailItem.EntryID ?? string.Empty;

                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                {
                    try
                    {
                        var id = $"{entryId}:{attachment.Position}:{attachment.FileName}";
                        attachments.Add(new AttachmentItem
                        {
                            Id = id,
                            FileName = attachment.FileName,
                            FileType = Path.GetExtension(attachment.FileName),
                            FileSizeBytes = attachment.Size,
                            SourceMailEntryId = entryId
                        });
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(attachment);
                    }
                }

                return attachments;
            }

            private static string GetInternetMessageId(Outlook.MailItem mailItem)
            {
                Outlook.PropertyAccessor? accessor = null;

                try
                {
                    accessor = mailItem.PropertyAccessor;
                    if (accessor is null)
                    {
                        return string.Empty;
                    }

                    var value = accessor.GetProperty(InternetMessageIdProperty);
                    return value as string ?? string.Empty;
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"GetInternetMessageId failed Entry='{mailItem.EntryID}' Property='{InternetMessageIdProperty}': {ex.Message}");
                    return string.Empty;
                }
                finally
                {
                    if (accessor is not null)
                    {
                        Marshal.ReleaseComObject(accessor);
                    }
                }
            }

            private static string GetThreadIndex(Outlook.MailItem mailItem)
            {
                Outlook.PropertyAccessor? accessor = null;

                try
                {
                    accessor = mailItem.PropertyAccessor;
                    if (accessor is null)
                    {
                        return string.Empty;
                    }

                    var value = accessor.GetProperty(ThreadIndexProperty);
                    if (value is byte[] buffer && buffer.Length > 0)
                    {
                        return Convert.ToBase64String(buffer);
                    }

                    if (value is string asString)
                    {
                        return asString;
                    }

                    return string.Empty;
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"GetThreadIndex failed Entry='{mailItem.EntryID}' Property='{ThreadIndexProperty}': {ex.Message}");
                    return string.Empty;
                }
                finally
                {
                    if (accessor is not null)
                    {
                        Marshal.ReleaseComObject(accessor);
                    }
                }
            }

            private static string GetThreadIndexPrefix(Outlook.MailItem mailItem, string threadIndex)
            {
                byte[]? indexBytes = null;
                Outlook.PropertyAccessor? accessor = null;

                try
                {
                    accessor = mailItem.PropertyAccessor;
                    if (accessor is not null)
                    {
                        var value = accessor.GetProperty(ConversationIndexProperty);
                        if (value is byte[] binary && binary.Length > 0)
                        {
                            indexBytes = binary;
                        }
                        else if (value is string raw && !string.IsNullOrWhiteSpace(raw))
                        {
                            indexBytes = DecodeThreadIndex(raw);
                        }
                    }
                }
                catch (COMException ex)
                {
                    DebugLogger.Log($"GetThreadIndexPrefix failed Entry='{mailItem.EntryID}' Property='{ConversationIndexProperty}': {ex.Message}");
                    indexBytes = null;
                }
                finally
                {
                    if (accessor is not null)
                    {
                        Marshal.ReleaseComObject(accessor);
                    }
                }

                if (indexBytes is null || indexBytes.Length == 0)
                {
                    indexBytes = DecodeThreadIndex(threadIndex);
                }

                return BuildThreadIndexPrefix(indexBytes);
            }

            private static byte[]? DecodeThreadIndex(string? encoded)
            {
                var trimmed = encoded?.Trim();
                if (string.IsNullOrEmpty(trimmed))
                {
                    return null;
                }

                try
                {
                    return Convert.FromBase64String(trimmed);
                }
                catch (FormatException)
                {
                    return null;
                }
            }

            private static string BuildThreadIndexPrefix(byte[]? indexBytes)
            {
                if (indexBytes is null || indexBytes.Length == 0)
                {
                    return string.Empty;
                }

                var length = Math.Min(indexBytes.Length, ThreadIndexPrefixBytes);
                return length > 0 ? Convert.ToBase64String(indexBytes, 0, length) : string.Empty;
            }

            private static IReadOnlyList<string> CollectReferenceMessageIds(Outlook.MailItem mailItem)
            {
                Outlook.PropertyAccessor? accessor = null;

                try
                {
                    accessor = mailItem.PropertyAccessor;
                    if (accessor is null)
                    {
                        return Array.Empty<string>();
                    }

                    var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    void AddFromProperty(string propertyName)
                    {
                        try
                        {
                            var propertyValue = accessor.GetProperty(propertyName) as string;
                            foreach (var id in ExtractMessageIds(propertyValue))
                            {
                                values.Add(id);
                            }
                        }
                        catch (COMException ex)
                        {
                            DebugLogger.Log($"CollectReferenceMessageIds failed Entry='{mailItem.EntryID}' Property='{propertyName}': {ex.Message}");
                        }
                    }

                    AddFromProperty(InReplyToProperty);
                    AddFromProperty(ReferencesProperty);

                    if (values.Count == 0)
                    {
                        try
                        {
                            var headers = accessor.GetProperty(TransportHeadersProperty) as string;
                            if (!string.IsNullOrEmpty(headers))
                            {
                                foreach (Match match in HeaderReferenceRegex.Matches(headers))
                                {
                                    var headerValue = match.Groups["value"].Value;
                                    foreach (var id in ExtractMessageIds(headerValue))
                                    {
                                        values.Add(id);
                                    }
                                }
                            }
                        }
                        catch (COMException ex)
                        {
                            DebugLogger.Log($"CollectReferenceMessageIds failed Entry='{mailItem.EntryID}' Property='{TransportHeadersProperty}': {ex.Message}");
                        }
                    }

                    return values.Count > 0 ? values.ToList() : Array.Empty<string>();
                }
                catch (COMException)
                {
                    return Array.Empty<string>();
                }
                finally
                {
                    if (accessor is not null)
                    {
                        Marshal.ReleaseComObject(accessor);
                    }
                }
            }

            private static DateTime? ToUniversal(DateTime value)
            {
                if (value == default)
                {
                    return null;
                }

                if (value.Kind == DateTimeKind.Utc)
                {
                    return value;
                }

                if (value.Kind == DateTimeKind.Local)
                {
                    return value.ToUniversalTime();
                }

                return DateTime.SpecifyKind(value, DateTimeKind.Local).ToUniversalTime();
            }

            private static IEnumerable<string> ExtractMessageIds(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    yield break;
                }

                var hasMatches = false;

                foreach (Match match in MessageIdRegex.Matches(value))
                {
                    hasMatches = true;
                    var id = match.Groups["id"].Value;
                    if (!string.IsNullOrWhiteSpace(id))
                    {
                        yield return id.Trim();
                    }
                }

                if (hasMatches)
                {
                    yield break;
                }

                // Fallback if no brackets found: split by common separators
                var parts = value!.Split(new[] { ' ', '\t', '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    var trimmed = part.Trim();
                    if (trimmed.Length == 0) continue;

                    if (trimmed.StartsWith("<", StringComparison.Ordinal) && trimmed.EndsWith(">", StringComparison.Ordinal) && trimmed.Length > 2)
                    {
                        trimmed = trimmed.Substring(1, trimmed.Length - 2).Trim();
                    }

                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        yield return trimmed;
                    }
                }
            }

            private static void ReleaseComObject<T>(ref T? comObject) where T : class
            {
                if (comObject is null)
                {
                    return;
                }

                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OutlookEventMonitor] Failed to release COM object: {ex}");
                }
                finally
                {
                    comObject = null;
                }
            }

            private void ThrowIfDisposed()
            {
                if (_isDisposed)
                {
                    throw new ObjectDisposedException(nameof(OutlookEventMonitor));
                }
            }

            private readonly struct ConversationCaptureStatus
            {
                public ConversationCaptureStatus(int added, int coverage, int totalCount, bool isComplete)
                {
                    Added = added;
                    Coverage = coverage;
                    TotalCount = totalCount;
                    IsComplete = isComplete;
                }

                public int Added { get; }

                public int Coverage { get; }

                public int TotalCount { get; }

                public bool IsComplete { get; }
            }

            private readonly struct ConversationEntry
            {
                public ConversationEntry(string entryId, string? storeId, DateTime? receivedUtc)
                {
                    EntryId = entryId ?? string.Empty;
                    StoreId = string.IsNullOrWhiteSpace(storeId) ? null : storeId;
                    ReceivedOnUtc = receivedUtc;
                }

                public string EntryId { get; }

                public string? StoreId { get; }

                public DateTime? ReceivedOnUtc { get; }
            }

            private sealed class ConversationEnumerationResult
            {
                public static readonly ConversationEnumerationResult Empty = new(Array.Empty<ConversationEntry>());

                public ConversationEnumerationResult(IReadOnlyList<ConversationEntry> entries)
                {
                    Entries = entries ?? Array.Empty<ConversationEntry>();
                }

                public IReadOnlyList<ConversationEntry> Entries { get; }

                public int TotalCount => Entries.Count;
            }

            private sealed class CatchUpMetadata
            {
                public IReadOnlyCollection<string>? StoreIds { get; set; }

                public IReadOnlyDictionary<string, IReadOnlyCollection<string>>? ConversationEntryIds { get; set; }

                public ISet<string>? ConversationTopics { get; set; }

                public ISet<string>? MessageIds { get; set; }

                public ISet<string>? NotFoundMessageIds { get; set; }

                public DateTime? EarliestReceivedOnUtc { get; set; }
            }

            public async Task<bool> TryProcessMailItemAsync(Outlook.MailItem mailItem, string? preferredEventId = null)
            {
                if (mailItem is null)
                {
                    return false;
                }

                try
                {
                    Outlook.MAPIFolder? parentFolder = null;
                    var snapshot = BuildSnapshot(mailItem, out parentFolder);
                    if (parentFolder != null)
                    {
                        Marshal.ReleaseComObject(parentFolder);
                    }

                    if (snapshot != null)
                    {
                        var result = await _eventRepository.TryAddMailAsync(snapshot, preferredEventId).ConfigureAwait(false);
                        return result != null;
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"TryProcessMailItemAsync failed: {ex.Message}");
                }

                return false;
            }

            public void RefreshCustomMonitors()
            {
                // Clean up existing custom monitors
                foreach (var items in _customFolderMonitors)
                {
                    items.ItemAdd -= OnCustomFolderItemAdded;
                    Marshal.ReleaseComObject(items);
                }
                _customFolderMonitors.Clear();

                // Re-initialize
                InitializeCustomMonitors();
            }

            private void InitializeCustomMonitors()
            {
                try
                {
                    if (Properties.Settings.Default.MonitoredFolders == null)
                    {
                        return;
                    }

                    var session = EnsureSession();
                    if (session == null) return;

                    foreach (string entryId in Properties.Settings.Default.MonitoredFolders)
                    {
                        if (string.IsNullOrWhiteSpace(entryId)) continue;

                        try
                        {
                            var folder = session.GetFolderFromID(entryId) as Outlook.Folder;
                            if (folder != null)
                            {
                                var items = folder.Items;
                                items.ItemAdd += OnCustomFolderItemAdded;
                                _customFolderMonitors.Add(items);
                                DebugLogger.Log($"Started monitoring custom folder: {folder.Name}");
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"Failed to monitor custom folder (EntryID: {entryId}): {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Error initializing custom folder monitors: {ex.Message}");
                }
            }

            private void OnCustomFolderItemAdded(object item)
            {
                if (!_isStarted) return;

                if (item is Outlook.MailItem mailItem)
                {
                    try
                    {
                        var entryId = mailItem.EntryID;
                        if (!string.IsNullOrWhiteSpace(entryId))
                        {
                            DebugLogger.Log($"OnCustomFolderItemAdded received EntryId='{entryId}' Subject='{mailItem.Subject}'");
                            _ = Task.Run(() => ProcessEntryIdsAsync(entryId));
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"OnCustomFolderItemAdded failed: {ex.Message}");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
                else
                {
                    Marshal.ReleaseComObject(item);
                }
            }
        }
}
