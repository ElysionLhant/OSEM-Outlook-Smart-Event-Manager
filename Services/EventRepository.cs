#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class EventRepository : IEventRepository, IDisposable
    {
        private readonly Application _application;
        private readonly string _storePath;
        private readonly SemaphoreSlim _gate = new(1, 1);
        private readonly List<EventRecord> _events = new();
        private readonly JsonSerializerSettings _serializerSettings;
        private readonly SynchronizationContext? _syncContext;
        private bool _isDisposed;
    private const string InternetMessageIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
    private const string ThreadIndexProperty = "http://schemas.microsoft.com/mapi/proptag/0x00710102";
    private const string ConversationIndexProperty = "http://schemas.microsoft.com/mapi/proptag/0x7F101102";
    private const string InReplyToProperty = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
    private const string ReferencesProperty = "http://schemas.microsoft.com/mapi/proptag/0x1039001F";
    private const string TransportHeadersProperty = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
    private static readonly TimeSpan DeduplicationWindow = TimeSpan.FromSeconds(30);
    private static readonly Regex HeaderReferenceRegex = new("(?im)^(?:References|In-Reply-To):\\s*(?<value>.+)$", RegexOptions.Compiled);
    private static readonly Regex MessageIdRegex = new("<(?<id>[^>]+)>", RegexOptions.Compiled);
    // Updated regex to support:
    // 1. Korean (제목) and Japanese (件名) headers
    // 2. Multiline/Folded headers (lines starting with whitespace)
    private static readonly Regex HistoricalSubjectRegex = new Regex(@Properties.Resources.Subject_s_s_subject_r_n_t, RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Multiline);
    private const int ThreadIndexPrefixBytes = 27; // First 27 bytes anchor the conversation root
    private const double ConversationTrackedWeight = 60d;
    private const double ReferenceMatchWeight = 70d;
    private const double ThreadPrefixReferenceWeight = 60d;
    private const double ThreadPrefixFingerprintWeight = 50d;
    private const double ThreadPrefixBaselineWeight = 45d;
    private const double ThreadHintFingerprintWeight = 40d;
    private const double ThreadHintBaselineWeight = 35d;
    private const double ThreadRootFingerprintWeight = 30d;
    private const double ThreadRootBaselineWeight = 25d;
    private const double SubjectParticipantsWeight = 35d;
    private const double SubjectMatchWeight = 50d;
    private const double ParticipantMatchWeight = 20d;
    private const double PreferredBiasWeight = 40d;
    private const double MinimumCandidateScore = 25d;

    private static readonly string[] SubjectPrefixes = new[] { 
        "RE:", "FW:", "FWD:", Properties.Resources.RE, Properties.Resources.FW, Properties.Resources.RE_1, Properties.Resources.FW_1, "[Pre-Alert]", 
        Properties.Resources.RE_2, Properties.Resources.FW_2, Properties.Resources.RE_3, Properties.Resources.FW_3,
        "[External]", "[EXT]", "Aw:", "Sv:", "Vs:" 
    };

    private static IReadOnlyList<string> ExtractHistoricalSubjects(string? body)
    {
        var subjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(body))
        {
            return Array.Empty<string>();
        }

        // 1. Try standard extraction
        var matches = HistoricalSubjectRegex.Matches(body!);
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
            // We validate the repair by checking if the repaired string contains a known header pattern
            if (EncodingRepair.TryFix(body!, s => HistoricalSubjectRegex.IsMatch(s), out var repairedBody))
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

    private static string NormalizeSubject(string? subject)
    {
        if (string.IsNullOrWhiteSpace(subject)) return string.Empty;
        
        // 1. NFKC Normalization (Full-width to Half-width)
        var normalized = subject!.Normalize(System.Text.NormalizationForm.FormKC);

        // Collapse multiple whitespace characters into a single space to handle formatting differences
        // e.g. "PO : 123" vs "PO :  123"
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized.Trim(), @"\s+", " ");
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

            if (!changed)
            {
                // Adaptive Mojibake Repair using the shared middleware
                // Validator: The repaired string must start with one of our known prefixes
                if (EncodingRepair.TryFix(normalized, 
                    candidate => SubjectPrefixes.Any(p => candidate.StartsWith(p, StringComparison.OrdinalIgnoreCase)), 
                    out var repaired))
                {
                    normalized = repaired;
                    changed = true;
                }
            }
        }
        return normalized;
    }

        public EventRepository(Application application, string? storePath = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _storePath = storePath ?? BuildDefaultStorePath();
            _syncContext = SynchronizationContext.Current;
            _serializerSettings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                TypeNameHandling = TypeNameHandling.None,
                DateTimeZoneHandling = DateTimeZoneHandling.Utc
            };
            Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
            LoadFromDisk();
        }

        public event EventHandler<EventChangedEventArgs>? EventChanged;

        public async Task<IReadOnlyList<EventRecord>> GetAllAsync()
        {
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                return _events.Select(Clone).ToList();
            }
            finally
            {
                _gate.Release();
            }
        }

        public IReadOnlyList<EventRecord> GetAll()
        {
            _gate.Wait();
            try
            {
                return _events.Select(Clone).ToList();
            }
            finally
            {
                _gate.Release();
            }
        }

        public EventRecord? GetEvent(string eventId)
        {
            _gate.Wait();
            try
            {
                return _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase))?.Let(Clone);
            }
            finally
            {
                _gate.Release();
            }
        }

        public async Task<EventRecord?> GetByIdAsync(string eventId)
        {
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                return _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase))?.Let(Clone);
            }
            finally
            {
                _gate.Release();
            }
        }

        public async Task<EventRecord> CreateFromMailAsync(MailItem mailItem, string? dashboardTemplateId = null, IEnumerable<string>? knownParticipants = null)
        {
            if (mailItem is null)
            {
                throw new ArgumentNullException(nameof(mailItem));
            }

            // Capture body on UI thread to avoid COM overhead in background task
            string mailBody = MailBodyFingerprint.GetBodyText(mailItem);

            // Run heavy analysis in background
            var (historicalSubjects, fingerprint) = await Task.Run(() => 
            {
                var subjects = ExtractHistoricalSubjects(mailBody);
                var fp = MailBodyFingerprint.Capture(mailBody);
                return (subjects, fp);
            });

            var record = new EventRecord
            {
                EventId = EventIdGenerator.Generate(),
                EventTitle = mailItem.Subject ?? Properties.Resources.No_Subject_Event,
                DashboardTemplateId = dashboardTemplateId ?? string.Empty,
                CreatedOn = DateTime.UtcNow,
                LastUpdatedOn = DateTime.UtcNow,
                ConversationIds = BuildConversationSet(mailItem)
            };

            // Initialize subject and participant collections
            var normalizedSubject = NormalizeSubject(mailItem.Subject);
            if (!string.IsNullOrEmpty(normalizedSubject))
            {
                record.RelatedSubjects.Add(normalizedSubject);
            }
            
            // Use pre-calculated historical subjects
            foreach (var hs in historicalSubjects)
            {
                var norm = NormalizeSubject(hs);
                if (!string.IsNullOrEmpty(norm))
                {
                    record.RelatedSubjects.Add(norm);
                }
            }

            var participants = knownParticipants?.ToList() ?? MailParticipantExtractor.Capture(mailItem).ToList();
            foreach (var p in participants)
            {
                record.Participants.Add(p);
            }

            record.Emails.Add(MapEmail(mailItem, isNewOrUpdated: true, precalculatedFingerprint: fingerprint, knownParticipants: participants));
            AppendAttachments(record, CaptureAttachments(mailItem));

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                _events.Add(record);
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }

            RaiseChanged(record, "Created");
            return Clone(record);
        }

        public async Task UpdateAsync(EventRecord record)
        {
            if (record is null)
            {
                throw new ArgumentNullException(nameof(record));
            }

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                var existing = _events.FirstOrDefault(e => string.Equals(e.EventId, record.EventId, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    throw new InvalidOperationException(string.Format(Properties.Resources.Event_record_EventId_does_not_exist, record.EventId));
                }

                CopyInto(record, existing);
                existing.LastUpdatedOn = DateTime.UtcNow;
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }

            RaiseChanged(record, "Updated");
        }

        public async Task ImportAsync(EventRecord record)
        {
            if (record is null) throw new ArgumentNullException(nameof(record));

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                var existing = _events.FirstOrDefault(e => string.Equals(e.EventId, record.EventId, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    _events.Add(Clone(record));
                }
                else
                {
                    CopyInto(record, existing);
                }
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }
            RaiseChanged(record, "Imported");
        }

        public async Task ArchiveAsync(IEnumerable<string> eventIds)
        {
            if (eventIds is null)
            {
                throw new ArgumentNullException(nameof(eventIds));
            }

            var normalized = eventIds.Where(id => !string.IsNullOrWhiteSpace(id)).Select(id => id.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (normalized.Count == 0)
            {
                return;
            }

            List<EventRecord> affected;
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                affected = _events.Where(e => normalized.Contains(e.EventId)).ToList();
                foreach (var record in affected)
                {
                    record.Status = EventStatus.Archived;
                    record.LastUpdatedOn = DateTime.UtcNow;
                }

                if (affected.Count > 0)
                {
                    await PersistLockedAsync().ConfigureAwait(false);
                }
            }
            finally
            {
                _gate.Release();
            }

            foreach (var record in affected)
            {
                RaiseChanged(Clone(record), "Archived");
            }
        }

        public async Task ReopenAsync(string eventId)
        {
            if (string.IsNullOrWhiteSpace(eventId))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(eventId));
            }

            EventRecord? record;
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                record = _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase));
                if (record is null)
                {
                    return;
                }

                record.Status = EventStatus.Open;
                record.LastUpdatedOn = DateTime.UtcNow;
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }

            if (record is not null)
            {
                RaiseChanged(Clone(record), "Reopened");
            }
        }

        public Task DeleteAsync(string eventId)
        {
            if (string.IsNullOrWhiteSpace(eventId))
            {
                return Task.CompletedTask;
            }

            return DeleteAsync(new[] { eventId });
        }

        public async Task DeleteAsync(IEnumerable<string> eventIds)
        {
            if (eventIds is null)
            {
                throw new ArgumentNullException(nameof(eventIds));
            }

            var normalized = eventIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (normalized.Count == 0)
            {
                return;
            }

            List<EventRecord> removed;

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                removed = _events.Where(e => normalized.Contains(e.EventId)).Select(Clone).ToList();
                if (removed.Count == 0)
                {
                    return;
                }

                _events.RemoveAll(e => normalized.Contains(e.EventId));
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }

            foreach (var record in removed)
            {
                RaiseChanged(record, "Deleted");
            }
        }

        async Task<EventRecord?> IEventRepository.TryAddMailAsync(MailSnapshot snapshot, string? preferredEventId)
        {
            if (snapshot is null || string.IsNullOrWhiteSpace(snapshot.ConversationId))
            {
                return null;
            }

            DebugLogger.Log($"TryAddMailAsync received EntryId='{snapshot.EntryId}', MsgId='{snapshot.InternetMessageId}', Conversation='{snapshot.ConversationId}', Store='{snapshot.StoreId}', Received='{snapshot.ReceivedOn:O}'");

            var emailModel = MapEmail(snapshot, isNewOrUpdated: true);
            var attachmentSnapshots = CloneAttachments(snapshot.Attachments);
            EventRecord? recordSnapshot = null;
            var changed = false;
            EventMatchCandidate? candidate = null;
            List<string> evaluationDiagnostics = new();

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                candidate = SelectEventCandidate(snapshot, preferredEventId, out evaluationDiagnostics);

                if (candidate is not null)
                {
                    var target = candidate.Record;

                    var previousCount = target.Emails.Count;

                    changed = UpsertMail(target, emailModel, attachmentSnapshots, allowRestore: false);
                    changed |= EnsureConversationTracked(target, snapshot.ConversationId);

                    // Also add historical subjects from the snapshot to RelatedSubjects
                    if (snapshot.HistoricalSubjects != null)
                    {
                        foreach (var hs in snapshot.HistoricalSubjects)
                        {
                            var norm = NormalizeSubject(hs);
                            if (!string.IsNullOrEmpty(norm))
                            {
                                if (target.RelatedSubjects == null)
                                {
                                    target.RelatedSubjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                }
                                if (target.RelatedSubjects.Add(norm))
                                {
                                    changed = true;
                                }
                            }
                        }
                    }

                    var currentCount = target.Emails.Count;
                    var hasNewContent = currentCount > previousCount;

                    if (changed)
                    {
                        // Fix: Update timestamp if content changed or was restored, not just if count increased.
                        // This ensures events jump to top when an email is un-removed or significantly updated.
                        target.LastUpdatedOn = DateTime.UtcNow;
                        await PersistLockedAsync().ConfigureAwait(false);
                    }

                    recordSnapshot = Clone(target);
                    candidate.HasNewContent = hasNewContent;
                }
            }
            finally
            {
                _gate.Release();
            }

            if (candidate is null)
            {
                var diag = evaluationDiagnostics.Count > 0 ? string.Join("; ", evaluationDiagnostics) : "(no diagnostics)";
                DebugLogger.Log($"TryAddMailAsync skipped append for EntryId='{emailModel.EntryId}' in Conversation='{emailModel.ConversationId}' candidate missing diagnostics={diag}");
                return null;
            }

            if (changed && recordSnapshot is not null)
            {
                var reason = candidate.ReasonSummary;
                var details = candidate.DetailSummary;
                DebugLogger.Log($"TryAddMailAsync processed mail EntryId='{emailModel.EntryId}' for Event='{recordSnapshot.EventId}' via {reason} details={details}");
                
                // Fix: Always raise changed event if data changed, even if count didn't increase.
                // This ensures UI refreshes for un-removals, moves, and metadata updates.
                if (candidate.HasNewContent)
                {
                    RaiseChanged(recordSnapshot, "MailAppended");
                }
                else
                {
                    RaiseChanged(recordSnapshot, "MailUpdated");
                }
            }
            else
            {
                var reason = candidate.ReasonSummary;
                var details = candidate.DetailSummary;
                DebugLogger.Log($"TryAddMailAsync skipped append for EntryId='{emailModel.EntryId}' in Conversation='{emailModel.ConversationId}' candidateScore={candidate.Score:F1} reasons={reason} details={details}");
            }

            return recordSnapshot;
        }

        public async Task<EventRecord?> AddMailToEventAsync(string eventId, MailItem mailItem)
        {
            if (string.IsNullOrWhiteSpace(eventId) || mailItem is null)
            {
                return null;
            }

            var emailModel = MapEmail(mailItem, isNewOrUpdated: true);
            var attachmentSnapshots = CaptureAttachments(mailItem);
            var mailConversationId = emailModel.ConversationId;
            EventRecord? snapshot = null;
            var changed = false;

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                var record = _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase));
                if (record is null)
                {
                    return null;
                }

                changed = UpsertMail(record, emailModel, attachmentSnapshots, allowRestore: true);
                changed |= EnsureConversationTracked(record, mailConversationId);

                // Update RelatedSubjects and Participants when manually adding mail
                var normalizedSubject = NormalizeSubject(mailItem.Subject);
                if (!string.IsNullOrEmpty(normalizedSubject))
                {
                    record.RelatedSubjects.Add(normalizedSubject);
                }

                // Extract and add historical subjects from body
                var historicalSubjects = ExtractHistoricalSubjects(mailItem.Body);
                foreach (var hs in historicalSubjects)
                {
                    var norm = NormalizeSubject(hs);
                    if (!string.IsNullOrEmpty(norm))
                    {
                        if (record.RelatedSubjects.Add(norm))
                        {
                            changed = true;
                        }
                    }
                }
                
                var participants = MailParticipantExtractor.Capture(mailItem);
                foreach (var p in participants)
                {
                    record.Participants.Add(p);
                }

                if (changed)
                {
                    record.LastUpdatedOn = DateTime.UtcNow;
                    await PersistLockedAsync().ConfigureAwait(false);
                }

                snapshot = Clone(record);
            }
            finally
            {
                _gate.Release();
            }

            if (changed && snapshot is not null)
            {
                RaiseChanged(snapshot, "MailAppended");
            }

            return snapshot;
        }

        public async Task RemoveMailAsync(string eventId, string entryId, string? internetMessageId = null)
        {
            if (string.IsNullOrWhiteSpace(eventId))
            {
                return;
            }

            EventRecord? record = null;
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                record = _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase));
                if (record is null)
                {
                    return;
                }

                EmailItem? email = null;
                if (!string.IsNullOrWhiteSpace(entryId))
                {
                    email = record.Emails.FirstOrDefault(e => string.Equals(e.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
                }

                if (email is null && !string.IsNullOrWhiteSpace(internetMessageId))
                {
                    email = record.Emails.FirstOrDefault(e => string.Equals(e.InternetMessageId, internetMessageId, StringComparison.OrdinalIgnoreCase));
                }

                if (email is null)
                {
                    DebugLogger.Log($"RemoveMailAsync unable to locate mail Event='{eventId}' Entry='{entryId}' MsgId='{internetMessageId}'");
                    return;
                }

                if (email.IsRemoved)
                {
                    return;
                }

                var targetEntryId = email.EntryId;
                email.IsRemoved = true;
                
                // Also remove the subject from RelatedSubjects if no other email shares it
                var subjectToRemove = NormalizeSubject(email.Subject);
                if (!string.IsNullOrEmpty(subjectToRemove) && record.RelatedSubjects != null)
                {
                    var isSubjectUsed = record.Emails
                        .Where(e => !e.IsRemoved && !string.Equals(e.EntryId, targetEntryId, StringComparison.OrdinalIgnoreCase))
                        .Any(e => string.Equals(NormalizeSubject(e.Subject), subjectToRemove, StringComparison.OrdinalIgnoreCase));
                    
                    if (!isSubjectUsed)
                    {
                        record.RelatedSubjects.Remove(subjectToRemove);
                    }
                }

                record.Attachments.RemoveAll(a => string.Equals(a.SourceMailEntryId, targetEntryId, StringComparison.OrdinalIgnoreCase));
                record.LastUpdatedOn = DateTime.UtcNow;
                await PersistLockedAsync().ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }

            if (record is not null)
            {
                DebugLogger.Log($"RemoveMailAsync marked mail Entry='{entryId}' MsgId='{internetMessageId}' as removed in Event='{record.EventId}'");
                RaiseChanged(Clone(record), "MailRemoved");
            }
        }

        public async Task MarkMessageIdsAsNotFoundAsync(string eventId, IEnumerable<string> messageIds)
        {
            if (string.IsNullOrWhiteSpace(eventId) || messageIds is null)
            {
                return;
            }

            var idsToAdd = messageIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (idsToAdd.Count == 0)
            {
                return;
            }

            EventRecord? record;
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                record = _events.FirstOrDefault(e => string.Equals(e.EventId, eventId, StringComparison.OrdinalIgnoreCase));
                if (record is null)
                {
                    return;
                }

                var changed = false;
                if (record.NotFoundMessageIds is null)
                {
                    record.NotFoundMessageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                foreach (var id in idsToAdd)
                {
                    if (record.NotFoundMessageIds.Add(id))
                    {
                        changed = true;
                    }
                }

                if (changed)
                {
                    record.LastUpdatedOn = DateTime.UtcNow;
                    await PersistLockedAsync().ConfigureAwait(false);
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _gate.Dispose();
            _isDisposed = true;
        }

        private static List<string> BuildConversationSet(MailItem mailItem)
        {
            var conversationId = mailItem.ConversationID;
            return string.IsNullOrWhiteSpace(conversationId) ? new List<string>() : new List<string> { conversationId };
        }

        private static EmailItem MapEmail(MailItem mailItem, bool isNewOrUpdated, string? precalculatedFingerprint = null, IEnumerable<string>? knownParticipants = null)
        {
            var participants = knownParticipants?.ToArray() ?? MailParticipantExtractor.Capture(mailItem);
            var referenceIds = CollectReferenceMessageIds(mailItem);
            var normalizedMessageId = NormalizeMessageId(GetInternetMessageId(mailItem));
            var threadIndex = GetThreadIndex(mailItem);
            var threadIndexPrefix = GetThreadIndexPrefix(mailItem, threadIndex);

            return new EmailItem
            {
                EntryId = mailItem.EntryID ?? Guid.NewGuid().ToString("N"),
                StoreId = GetStoreId(mailItem),
                ConversationId = mailItem.ConversationID ?? string.Empty,
                InternetMessageId = normalizedMessageId,
                Sender = mailItem.SenderName ?? string.Empty,
                To = mailItem.To ?? string.Empty,
                Subject = mailItem.Subject ?? string.Empty,
                Participants = participants,
                BodyFingerprint = precalculatedFingerprint ?? MailBodyFingerprint.Capture(mailItem),
                ThreadIndex = threadIndex,
                ThreadIndexPrefix = threadIndexPrefix,
                ReferenceMessageIds = NormalizeMessageIds(referenceIds),
                ReceivedOn = mailItem.ReceivedTime.ToUniversalTime(),
                IsNewOrUpdated = isNewOrUpdated,
                IsRemoved = false
            };

        }

        private static EmailItem MapEmail(MailSnapshot snapshot, bool isNewOrUpdated)
        {
            var participants = snapshot.Participants is string[] participantArray
                ? participantArray
                : snapshot.Participants?.ToArray() ?? Array.Empty<string>();
            var normalizedMessageId = NormalizeMessageId(snapshot.InternetMessageId);
            var normalizedReferences = snapshot.ReferenceMessageIds is null
                ? Array.Empty<string>()
                : NormalizeMessageIds(snapshot.ReferenceMessageIds);

            return new EmailItem
            {
                EntryId = string.IsNullOrEmpty(snapshot.EntryId) ? Guid.NewGuid().ToString("N") : snapshot.EntryId,
                StoreId = snapshot.StoreId ?? string.Empty,
                ConversationId = snapshot.ConversationId ?? string.Empty,
                InternetMessageId = normalizedMessageId,
                Sender = snapshot.Sender ?? string.Empty,
                To = snapshot.To ?? string.Empty,
                Subject = snapshot.Subject ?? string.Empty,
                Participants = participants,
                BodyFingerprint = snapshot.BodyFingerprint ?? string.Empty,
                ThreadIndex = snapshot.ThreadIndex ?? string.Empty,
                ThreadIndexPrefix = snapshot.ThreadIndexPrefix ?? string.Empty,
                ReferenceMessageIds = normalizedReferences,
                ReceivedOn = snapshot.ReceivedOn,
                IsNewOrUpdated = isNewOrUpdated,
                IsRemoved = false
            };
        }

        private static List<AttachmentItem> CaptureAttachments(MailItem mailItem)
        {
            var entryId = mailItem.EntryID ?? string.Empty;
            var attachments = new List<AttachmentItem>();
            foreach (Attachment attachment in mailItem.Attachments)
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

        private static List<AttachmentItem> CloneAttachments(IEnumerable<AttachmentItem>? attachments)
        {
            if (attachments is null)
            {
                return new List<AttachmentItem>();
            }

            var clones = new List<AttachmentItem>();
            foreach (var attachment in attachments)
            {
                if (attachment is null)
                {
                    continue;
                }

                clones.Add(new AttachmentItem
                {
                    Id = attachment.Id,
                    FileName = attachment.FileName,
                    FileType = attachment.FileType,
                    FileSizeBytes = attachment.FileSizeBytes,
                    SourceMailEntryId = attachment.SourceMailEntryId
                });
            }

            return clones;
        }

        private static string GetStoreId(MailItem mailItem)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            MAPIFolder? folder = null;
            try
            {
                folder = mailItem.Parent as MAPIFolder;
                return folder?.StoreID ?? string.Empty;
            }
            finally
            {
                if (folder is not null)
                {
                    Marshal.ReleaseComObject(folder);
                }
            }
        }

        private static string GetInternetMessageId(MailItem mailItem)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            PropertyAccessor? accessor = null;
            try
            {
                accessor = mailItem.PropertyAccessor;
                var value = accessor?.GetProperty(InternetMessageIdProperty);
                return value as string ?? string.Empty;
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
                if (accessor is not null)
                {
                    Marshal.ReleaseComObject(accessor);
                }
            }
        }

        private static string GetThreadIndex(MailItem mailItem)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            PropertyAccessor? accessor = null;
            try
            {
                accessor = mailItem.PropertyAccessor;
                var value = accessor?.GetProperty(ThreadIndexProperty);
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
                if (accessor is not null)
                {
                    Marshal.ReleaseComObject(accessor);
                }
            }
        }

        private static string GetThreadIndexPrefix(MailItem mailItem, string threadIndex)
        {
            if (mailItem is null)
            {
                return string.Empty;
            }

            byte[]? indexBytes = null;
            PropertyAccessor? accessor = null;

            try
            {
                accessor = mailItem.PropertyAccessor;
                if (accessor is not null)
                {
                    var value = accessor.GetProperty(ConversationIndexProperty);
                    if (value is byte[] buffer && buffer.Length > 0)
                    {
                        indexBytes = buffer;
                    }
                    else if (value is string raw && !string.IsNullOrWhiteSpace(raw))
                    {
                        indexBytes = DecodeThreadIndex(raw);
                    }
                }
            }
            catch (COMException)
            {
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

        private static IReadOnlyList<string> CollectReferenceMessageIds(MailItem mailItem)
        {
            if (mailItem is null)
            {
                return Array.Empty<string>();
            }

            PropertyAccessor? accessor = null;
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
                    catch (COMException)
                    {
                        // ignore individual property read failures
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
                    catch (COMException)
                    {
                        // ignore header read failures
                    }
                }

                return values.Count > 0 ? values.ToArray() : Array.Empty<string>();
            }
            catch (COMException)
            {
                return Array.Empty<string>();
            }
            catch
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

        private static IEnumerable<string> ExtractMessageIds(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                yield break;
            }

            var hasMatches = false;

            var source = value!;

            foreach (Match match in MessageIdRegex.Matches(source))
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

            var trimmed = source.Trim();
            if (trimmed.Length == 0)
            {
                yield break;
            }

            if (trimmed.StartsWith("<", StringComparison.Ordinal) && trimmed.EndsWith(">", StringComparison.Ordinal) && trimmed.Length > 2)
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2).Trim();
            }

            if (!string.IsNullOrEmpty(trimmed))
            {
                yield return trimmed;
            }
        }

        private static bool UpsertMail(EventRecord record, EmailItem email, List<AttachmentItem> attachments, bool allowRestore)
        {
            if (record is null || email is null)
            {
                return false;
            }

            // Ensure RelatedSubjects is updated with the new email's subject
            var normalizedSubject = NormalizeSubject(email.Subject);
            if (!string.IsNullOrEmpty(normalizedSubject))
            {
                if (record.RelatedSubjects == null)
                {
                    record.RelatedSubjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }
                record.RelatedSubjects.Add(normalizedSubject);
            }

            var existing = FindExistingEmail(record, email);
            if (existing is null)
            {
                record.Emails.Add(email);
                AppendAttachments(record, attachments);
                return true;
            }

            if (existing.IsRemoved && !allowRestore)
            {
                DebugLogger.Log($"UpsertMail skipped previously removed mail Entry='{existing.EntryId}' MsgId='{existing.InternetMessageId}' Conversation='{existing.ConversationId}'");
                return false;
            }

            // Highlight Logic: Check if this message was previously processed (highlighted) and cleared
            if (record.ProcessedMessageIds != null && !string.IsNullOrEmpty(email.InternetMessageId))
            {
                if (record.ProcessedMessageIds.Contains(email.InternetMessageId))
                {
                    // If it was processed before, ensure we don't mark it as new again
                    email.IsNewOrUpdated = false;
                }
            }

            var changed = false;
            var contentChanged = false;
            var originalEntryId = existing.EntryId;

            if (!string.IsNullOrEmpty(email.EntryId) && !string.Equals(existing.EntryId, email.EntryId, StringComparison.OrdinalIgnoreCase))
            {
                existing.EntryId = email.EntryId;
                changed = true;
            }

            if (!string.IsNullOrEmpty(email.StoreId) && !string.Equals(existing.StoreId, email.StoreId, StringComparison.OrdinalIgnoreCase))
            {
                existing.StoreId = email.StoreId;
                changed = true;
            }

            if (!string.IsNullOrEmpty(email.InternetMessageId) && !string.Equals(existing.InternetMessageId, email.InternetMessageId, StringComparison.OrdinalIgnoreCase))
            {
                existing.InternetMessageId = email.InternetMessageId;
                changed = true;
            }

            if (!string.Equals(existing.Sender, email.Sender, StringComparison.OrdinalIgnoreCase))
            {
                existing.Sender = email.Sender;
                changed = true;
                contentChanged = true;
            }

            if (!string.Equals(existing.Subject, email.Subject, StringComparison.OrdinalIgnoreCase))
            {
                existing.Subject = email.Subject;
                changed = true;
                contentChanged = true;
            }

            if (existing.ReceivedOn != email.ReceivedOn)
            {
                existing.ReceivedOn = email.ReceivedOn;
                changed = true;
            }

            if (!string.IsNullOrEmpty(email.BodyFingerprint) && !string.Equals(existing.BodyFingerprint, email.BodyFingerprint, StringComparison.Ordinal))
            {
                existing.BodyFingerprint = email.BodyFingerprint;
                changed = true;
                contentChanged = true;
            }

            if (!string.IsNullOrEmpty(email.ThreadIndex) && !string.Equals(existing.ThreadIndex, email.ThreadIndex, StringComparison.Ordinal))
            {
                existing.ThreadIndex = email.ThreadIndex;
                changed = true;
            }

            if (!string.IsNullOrEmpty(email.ThreadIndexPrefix) && !string.Equals(existing.ThreadIndexPrefix, email.ThreadIndexPrefix, StringComparison.Ordinal))
            {
                existing.ThreadIndexPrefix = email.ThreadIndexPrefix;
                changed = true;
            }

            if (email.ReferenceMessageIds is not null && email.ReferenceMessageIds.Length > 0)
            {
                var existingReferences = existing.ReferenceMessageIds ?? Array.Empty<string>();
                var referenceSet = new HashSet<string>(existingReferences, StringComparer.OrdinalIgnoreCase);
                var addedReference = false;
                foreach (var referenceId in email.ReferenceMessageIds)
                {
                    if (string.IsNullOrWhiteSpace(referenceId))
                    {
                        continue;
                    }

                    if (referenceSet.Add(referenceId.Trim()))
                    {
                        addedReference = true;
                    }
                }

                if (addedReference)
                {
                    existing.ReferenceMessageIds = referenceSet.ToArray();
                    changed = true;
                }
            }

            if (email.Participants is not null && email.Participants.Length > 0)
            {
                var existingParticipants = existing.Participants ?? Array.Empty<string>();
                
                // Fix: Use SetEquals to ignore order differences which caused false 'contentChanged'
                var existingSet = new HashSet<string>(existingParticipants, StringComparer.OrdinalIgnoreCase);
                var newSet = new HashSet<string>(email.Participants, StringComparer.OrdinalIgnoreCase);
                
                if (!existingSet.SetEquals(newSet))
                {
                    existing.Participants = email.Participants;
                    changed = true;
                    contentChanged = true;
                }
            }

            if (existing.IsRemoved)
            {
                existing.IsRemoved = false;
                changed = true;
                contentChanged = true;
            }

            // Only mark as updated if something actually changed, AND it wasn't already processed
            // If it's just a refresh finding the same email, we shouldn't re-highlight it unless data changed significantly
            // However, the logic above sets 'changed = true' for many properties.
            // If 'changed' is true, we might want to highlight it.
            // But user complains about "refresh highlights again".
            // This implies that even if nothing changed, or maybe trivial changes, it gets highlighted.
            // Or maybe 'changed' is false, but we force it?
            
            // The previous code block was:
            // if (!existing.IsNewOrUpdated)
            // {
            //     existing.IsNewOrUpdated = true;
            //     changed = true;
            // }
            // else
            // {
            //     existing.IsNewOrUpdated = true;
            // }
            
            // This unconditionally sets IsNewOrUpdated = true whenever UpsertMail is called for an existing mail!
            // This is the bug. UpsertMail is called during CatchUp (Refresh).
            
            if (changed)
            {
                 // Fix: Check ProcessedMessageIds before re-highlighting
                 var isProcessed = record.ProcessedMessageIds != null && 
                                   !string.IsNullOrEmpty(existing.InternetMessageId) && 
                                   record.ProcessedMessageIds.Contains(existing.InternetMessageId);

                 if (!existing.IsNewOrUpdated && contentChanged && !isProcessed)
                 {
                     existing.IsNewOrUpdated = true;
                 }
            }

            var removed = record.Attachments.RemoveAll(a => string.Equals(a.SourceMailEntryId, originalEntryId, StringComparison.OrdinalIgnoreCase));
            if (removed > 0)
            {
                changed = true;
            }

            if (!string.Equals(originalEntryId, email.EntryId, StringComparison.OrdinalIgnoreCase))
            {
                removed = record.Attachments.RemoveAll(a => string.Equals(a.SourceMailEntryId, email.EntryId, StringComparison.OrdinalIgnoreCase));
                if (removed > 0)
                {
                    changed = true;
                }
            }

            AppendAttachments(record, attachments);
            if (attachments.Count > 0)
            {
                changed = true;
            }

            return changed;
        }

        private static EmailItem? FindExistingEmail(EventRecord record, EmailItem candidate)
        {
            return record.Emails.FirstOrDefault(email => IsSameMail(email, candidate));
        }

        private EventMatchCandidate? SelectEventCandidate(MailSnapshot snapshot, string? preferredEventId, out List<string> diagnostics)
        {
            diagnostics = new List<string>();

            if (snapshot is null)
            {
                diagnostics.Add("snapshot null");
                return null;
            }

            var openEvents = _events.Where(e => e.Status == EventStatus.Open).ToList();
            DebugLogger.Log($"SelectEventCandidate invoked EntryId='{snapshot.EntryId}' MsgId='{snapshot.InternetMessageId}' Conversation='{snapshot.ConversationId}' references='{snapshot.ReferenceMessageIds?.Count ?? 0}'");

            if (openEvents.Count == 0)
            {
                diagnostics.Add("no open events available");
                return null;
            }

            var candidates = new Dictionary<string, EventMatchCandidate>(StringComparer.OrdinalIgnoreCase);

            EventMatchCandidate GetOrCreateCandidate(EventRecord record)
            {
                if (!candidates.TryGetValue(record.EventId, out var candidate))
                {
                    candidate = new EventMatchCandidate(record);
                    candidates[record.EventId] = candidate;
                }

                return candidate;
            }

            void AddCandidate(EventRecord record, double weight, string reason, string detail)
            {
                var candidate = GetOrCreateCandidate(record);
                candidate.AddScore(weight, reason, detail);
            }

            var conversationId = snapshot.ConversationId?.Trim();
            // Disabled per user requirement: Do not use ConversationID for ingestion. Rely on Subject + Participants.
            /*
            if (!string.IsNullOrEmpty(conversationId))
            {
                var conversationMatches = 0;
                var conversationKey = conversationId!;
                foreach (var record in openEvents)
                {
                    if (record.IsConversationTracked(conversationKey))
                    {
                        AddCandidate(record, ConversationTrackedWeight, "conversation", $"tracked conversation '{conversationKey}'");
                        conversationMatches++;
                    }
                }

                if (conversationMatches == 0)
                {
                    diagnostics.Add($"conversationId '{conversationKey}' not tracked by any open event");
                }
            }
            else
            {
                diagnostics.Add("conversationId missing");
            }
            */

            /*
            HashSet<string>? referenceSet = null;
            if (snapshot.ReferenceMessageIds is not null && snapshot.ReferenceMessageIds.Count > 0)
            {
                referenceSet = new HashSet<string>(
                    snapshot.ReferenceMessageIds
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Select(NormalizeMessageId)
                        .Where(id => !string.IsNullOrEmpty(id)),
                    StringComparer.OrdinalIgnoreCase);

                if (referenceSet.Count > 0)
                {
                    var referenceMatches = 0;
                    foreach (var record in openEvents)
                    {
                        foreach (var email in record.Emails)
                        {
                            if (string.IsNullOrEmpty(email.InternetMessageId))
                            {
                                continue;
                            }

                            var normalizedInternetMessageId = NormalizeMessageId(email.InternetMessageId);
                            if (!string.IsNullOrEmpty(normalizedInternetMessageId) && referenceSet.Contains(normalizedInternetMessageId))
                            {
                                AddCandidate(record, ReferenceMatchWeight, "reference", $"reference '{email.InternetMessageId}'");
                                referenceMatches++;
                                break;
                            }
                        }
                    }

                    if (referenceMatches == 0)
                    {
                        diagnostics.Add($"references examined={referenceSet.Count} no match");
                    }
                }
                else
                {
                    diagnostics.Add("references normalized=0");
                }
            }
            else
            {
                diagnostics.Add("references missing");
            }
            */

            /*
            var threadPrefix = snapshot.ThreadIndexPrefix;
            if (!string.IsNullOrEmpty(threadPrefix))
            {
                var attempts = new List<string>();
                var matches = 0;

                foreach (var record in openEvents)
                {
                    var eventMatched = false;

                    foreach (var email in record.Emails)
                    {
                        if (eventMatched)
                        {
                            break;
                        }

                        if (string.IsNullOrEmpty(email.ThreadIndexPrefix))
                        {
                            continue;
                        }

                        if (!string.Equals(threadPrefix, email.ThreadIndexPrefix, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        var referenceMatched = referenceSet is not null &&
                                               referenceSet.Count > 0 &&
                                               !string.IsNullOrEmpty(email.InternetMessageId) &&
                                               referenceSet.Contains(NormalizeMessageId(email.InternetMessageId));

                        var similarity = (!string.IsNullOrEmpty(snapshot.BodyFingerprint) && !string.IsNullOrEmpty(email.BodyFingerprint))
                            ? MailBodyFingerprint.ComputeSimilarityScore(snapshot.BodyFingerprint, email.BodyFingerprint)
                            : 0d;

                        var fingerprintMatched = similarity >= 0.6;
                        var baselineMatched = !string.IsNullOrEmpty(snapshot.BodyFingerprint) &&
                                              !string.IsNullOrEmpty(email.BodyFingerprint) &&
                                              MailBodyFingerprint.MatchesBaseline(snapshot.BodyFingerprint, email.BodyFingerprint);

                        var matchedSomething = false;

                        if (referenceMatched)
                        {
                            AddCandidate(record, ThreadPrefixReferenceWeight, "threadPrefix+reference", $"threadPrefix '{threadPrefix}' reference '{email.InternetMessageId}'");
                            matchedSomething = true;
                        }

                        if (fingerprintMatched)
                        {
                            AddCandidate(record, ThreadPrefixFingerprintWeight, "threadPrefix+fingerprint", $"threadPrefix '{threadPrefix}' similarity={similarity:F2}");
                            matchedSomething = true;
                        }

                        if (baselineMatched)
                        {
                            AddCandidate(record, ThreadPrefixBaselineWeight, "threadPrefix+baseline", $"threadPrefix '{threadPrefix}' baseline match");
                            matchedSomething = true;
                        }

                        if (matchedSomething)
                        {
                            matches++;
                            eventMatched = true;
                        }
                        else
                        {
                            attempts.Add($"Event={record.EventId} ref={referenceMatched} sim={similarity:F2} baseline={baselineMatched} bodyEmpty={string.IsNullOrEmpty(email.BodyFingerprint)}");
                        }
                    }
                }

                if (matches == 0)
                {
                    diagnostics.Add(attempts.Count > 0
                        ? $"threadPrefix attempts={attempts.Count} -> {SummarizeDiagnostics(attempts)}"
                        : $"threadPrefix '{threadPrefix}' produced no candidates");
                }
            }
            else
            {
                diagnostics.Add("threadPrefix missing");
            }
            */

            // New Logic: Subject and Participant Matching
            var normalizedSubject = NormalizeSubject(snapshot.Subject);
            var historicalSubjects = snapshot.HistoricalSubjects?
                .Select(NormalizeSubject)
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList() ?? new List<string>();

            if (!string.IsNullOrEmpty(normalizedSubject) || historicalSubjects.Count > 0)
            {
                var subjectMatches = 0;
                foreach (var record in openEvents)
                {
                    // Strict check: Only match if the subject matches the Event Title (normalized)
                    // We ignore RelatedSubjects because it may contain polluted subjects from previous incorrect merges.
                    // This ensures that emails from Client A don't get matched to Client B events just because they were wrongly merged once.
                    var normalizedTitle = NormalizeSubject(record.EventTitle);
                    
                    // Also check the first email's subject as a fallback if the title was renamed to something completely different
                    var firstEmailSubject = record.Emails.FirstOrDefault()?.Subject;
                    var normalizedFirstSubject = NormalizeSubject(firstEmailSubject);

                    // Helper to check for match (Equals or StartsWith)
                    bool IsStandardMatch(string? input, string? target)
                    {
                        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(target)) return false;
                        // Allow exact match OR prefix match (e.g. "Subject / Suffix" matches "Subject")
                        return string.Equals(input, target, StringComparison.OrdinalIgnoreCase) ||
                               input!.StartsWith(target!, StringComparison.OrdinalIgnoreCase);
                    }

                    // Helper to check for truncation match (Target starts with Input)
                    bool IsTruncatedMatch(string? input, string? target)
                    {
                        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(target)) return false;
                        // We require the input to be at least 4 chars long to avoid matching generic prefixes like "Re:"
                        return input!.Length >= 4 && target!.StartsWith(input, StringComparison.OrdinalIgnoreCase);
                    }

                    // 1. Check Standard Match (Exact or Prefix)
                    var isStandardMatch = IsStandardMatch(normalizedSubject, normalizedTitle) ||
                                          IsStandardMatch(normalizedSubject, normalizedFirstSubject) ||
                                          (record.RelatedSubjects != null && record.RelatedSubjects.Any(rs => IsStandardMatch(normalizedSubject, rs)));

                    // 2. Check Truncated Match (if standard failed)
                    var isTruncatedMatch = !isStandardMatch && (
                                           IsTruncatedMatch(normalizedSubject, normalizedTitle) ||
                                           IsTruncatedMatch(normalizedSubject, normalizedFirstSubject) ||
                                           (record.RelatedSubjects != null && record.RelatedSubjects.Any(rs => IsTruncatedMatch(normalizedSubject, rs))));

                    // 3. Check Historical Matches
                    // If we have a Truncated Match on the header, we REQUIRE a Standard Match in the history to confirm.
                    // If we have a Standard Match on the header, history is just a bonus.
                    var historicalMatch = false;
                    foreach (var histSubject in historicalSubjects)
                    {
                        if (IsStandardMatch(histSubject, normalizedTitle) ||
                            IsStandardMatch(histSubject, normalizedFirstSubject) ||
                            (record.RelatedSubjects != null && record.RelatedSubjects.Any(rs => IsStandardMatch(histSubject, rs))))
                        {
                            historicalMatch = true;
                            break;
                        }
                    }

                    // Final Decision
                    // - Standard Match: Accepted
                    // - Truncated Match: Accepted ONLY IF Historical Match is present (Double Confirmation)
                    // - Historical Match: Accepted (as per original logic, though usually implies header mismatch)
                    
                    if (isStandardMatch || (isTruncatedMatch && historicalMatch) || historicalMatch)
                    {
                        // Check participants
                        bool participantMatch = false;
                        if (snapshot.Participants != null && snapshot.Participants.Count > 0 && record.Participants != null && record.Participants.Count > 0)
                        {
                            // Check if any participant in the snapshot matches any participant in the record
                            // Or should it be stricter? "收发件人相同" -> Sender/Recipient match.
                            // The snapshot.Participants includes Sender, To, CC.
                            // We check for intersection.
                            if (MailParticipantExtractor.Intersects(snapshot.Participants, record.Participants))
                            {
                                participantMatch = true;
                            }
                        }

                        if (participantMatch)
                        {
                            var matchType = isStandardMatch ? "subject" : (isTruncatedMatch ? "truncated-subject+history" : "historical-subject");
                            AddCandidate(record, SubjectMatchWeight + ParticipantMatchWeight, $"{matchType}+participant", $"{matchType} matched participants matched");
                            subjectMatches++;
                        }
                        else
                        {
                            // Only subject matched, maybe give lower score or don't match at all?
                            // User said: "先搜索Subject后校验参与者，都通过再收入" -> Both must pass.
                            diagnostics.Add($"Event={record.EventId} subject matched but participants mismatch");
                        }
                    }
                }
                
                if (subjectMatches == 0)
                {
                    diagnostics.Add($"subject '{normalizedSubject}' (and {historicalSubjects.Count} historical) no full match");
                }
            }

            var snapshotFingerprint = snapshot.BodyFingerprint;

            /*
            var threadHint = GetThreadIndexHint(threadPrefix, snapshot.ThreadIndex);
            if (!string.IsNullOrEmpty(threadHint) && !string.IsNullOrEmpty(snapshotFingerprint))
            {
                var attempts = new List<string>();
                var matches = 0;

                foreach (var record in openEvents)
                {
                    var eventMatched = false;

                    foreach (var email in record.Emails)
                    {
                        if (eventMatched)
                        {
                            break;
                        }

                        if (string.IsNullOrEmpty(email.ThreadIndex) && string.IsNullOrEmpty(email.ThreadIndexPrefix))
                        {
                            continue;
                        }

                        var emailHint = GetThreadIndexHint(email.ThreadIndexPrefix, email.ThreadIndex);
                        if (string.IsNullOrEmpty(emailHint))
                        {
                            continue;
                        }

                        if (!string.Equals(threadHint, emailHint, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        var similarity = (!string.IsNullOrEmpty(email.BodyFingerprint))
                            ? MailBodyFingerprint.ComputeSimilarityScore(snapshotFingerprint!, email.BodyFingerprint)
                            : 0d;

                        var fingerprintMatched = similarity >= 0.6;
                        var baselineMatched = MailBodyFingerprint.MatchesBaseline(snapshotFingerprint!, email.BodyFingerprint);

                        if (fingerprintMatched)
                        {
                            AddCandidate(record, ThreadHintFingerprintWeight, "threadHint+fingerprint", $"threadHint '{threadHint}' similarity={similarity:F2}");
                            matches++;
                            eventMatched = true;
                        }
                        else if (baselineMatched)
                        {
                            AddCandidate(record, ThreadHintBaselineWeight, "threadHint+baseline", $"threadHint '{threadHint}' baseline match");
                            matches++;
                            eventMatched = true;
                        }
                        else
                        {
                            attempts.Add($"Event={record.EventId} sim={similarity:F2} baseline={baselineMatched} bodyEmpty={string.IsNullOrEmpty(email.BodyFingerprint)}");
                        }
                    }
                }

                if (matches == 0)
                {
                    diagnostics.Add(attempts.Count > 0
                        ? $"threadHint attempts={attempts.Count} -> {SummarizeDiagnostics(attempts)}"
                        : $"threadHint '{threadHint}' produced no candidates");
                }
            }
            else
            {
                if (string.IsNullOrEmpty(threadHint))
                {
                    diagnostics.Add("threadHint missing");
                }
                else
                {
                    diagnostics.Add("threadHint skipped due to empty fingerprint");
                }
            }

            var threadRoot = GetThreadRoot(snapshot.ThreadIndex);
            if (!string.IsNullOrEmpty(threadRoot) && !string.IsNullOrEmpty(snapshotFingerprint))
            {
                var attempts = new List<string>();
                var matches = 0;

                foreach (var record in openEvents)
                {
                    var eventMatched = false;

                    foreach (var email in record.Emails)
                    {
                        if (eventMatched)
                        {
                            break;
                        }

                        if (string.IsNullOrEmpty(email.ThreadIndex) || string.IsNullOrEmpty(email.BodyFingerprint))
                        {
                            continue;
                        }

                        var existingRoot = GetThreadRoot(email.ThreadIndex);
                        if (string.IsNullOrEmpty(existingRoot))
                        {
                            continue;
                        }

                        if (!string.Equals(threadRoot, existingRoot, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        var similarity = MailBodyFingerprint.ComputeSimilarityScore(snapshotFingerprint!, email.BodyFingerprint);
                        var fingerprintMatched = similarity >= 0.6;
                        var baselineMatched = MailBodyFingerprint.MatchesBaseline(snapshotFingerprint!, email.BodyFingerprint);

                        if (fingerprintMatched)
                        {
                            AddCandidate(record, ThreadRootFingerprintWeight, "threadRoot+fingerprint", $"threadRoot '{threadRoot}' similarity={similarity:F2}");
                            matches++;
                            eventMatched = true;
                        }
                        else if (baselineMatched)
                        {
                            AddCandidate(record, ThreadRootBaselineWeight, "threadRoot+baseline", $"threadRoot '{threadRoot}' baseline match");
                            matches++;
                            eventMatched = true;
                        }
                        else
                        {
                            attempts.Add($"Event={record.EventId} sim={similarity:F2} baseline={baselineMatched}");
                        }
                    }
                }

                if (matches == 0)
                {
                    diagnostics.Add(attempts.Count > 0
                        ? $"threadRoot attempts={attempts.Count} -> {SummarizeDiagnostics(attempts)}"
                        : $"threadRoot '{threadRoot}' produced no candidates");
                }
            }
            else
            {
                diagnostics.Add($"threadRoot skipped root='{threadRoot}' fingerprintEmpty={string.IsNullOrEmpty(snapshotFingerprint)}");
            }
            */

            /*
            var normalizedParticipants = MailParticipantExtractor.Normalize(snapshot.Participants);
            var subjectKey = NormalizeSubjectForFallback(snapshot.Subject);

            if (!string.IsNullOrEmpty(subjectKey) && normalizedParticipants.Length > 0)
            {
                var attempts = new List<string>();
                var matches = 0;

                foreach (var record in openEvents)
                {
                    var eventMatched = false;

                    foreach (var email in record.Emails)
                    {
                        if (eventMatched)
                        {
                            break;
                        }

                        if (string.IsNullOrEmpty(email.Subject))
                        {
                            continue;
                        }

                        var emailSubjectKey = NormalizeSubjectForFallback(email.Subject);
                        if (string.IsNullOrEmpty(emailSubjectKey) || !string.Equals(subjectKey, emailSubjectKey, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        var overlap = MailParticipantExtractor.ComputeOverlapScore(normalizedParticipants, email.Participants);
                        if (overlap >= 0.5)
                        {
                            AddCandidate(record, SubjectParticipantsWeight, "participants+subject", $"overlap={overlap:F2} subject='{subjectKey}'");
                            matches++;
                            eventMatched = true;
                        }
                        else
                        {
                            attempts.Add($"Event={record.EventId} overlap={overlap:F2}");
                        }
                    }
                }

                if (matches == 0)
                {
                    diagnostics.Add(attempts.Count > 0
                        ? $"participants+subject attempts={attempts.Count} -> {SummarizeDiagnostics(attempts)}"
                        : $"participants+subject subject='{subjectKey}' produced no candidates");
                }
            }
            else
            {
                diagnostics.Add($"participants+subject skipped subjectEmpty={string.IsNullOrEmpty(subjectKey)} participantCount={normalizedParticipants.Length}");
            }
            */

            var preferredKey = preferredEventId?.Trim();
            if (!string.IsNullOrEmpty(preferredKey))
            {
                if (candidates.TryGetValue(preferredKey!, out var preferredCandidate))
                {
                    preferredCandidate.ApplyPreferredBias(PreferredBiasWeight, $"preferredId '{preferredKey}'");
                }
                else
                {
                    diagnostics.Add($"preferred event '{preferredKey}' produced no candidate");
                }
            }

            EventMatchCandidate? best = null;
            foreach (var candidate in candidates.Values)
            {
                if (best is null || candidate.Score > best.Score + 0.01d)
                {
                    best = candidate;
                    continue;
                }

                if (Math.Abs(candidate.Score - best.Score) <= 0.01d)
                {
                    if (candidate.PreferredBiasApplied && !best.PreferredBiasApplied)
                    {
                        best = candidate;
                        continue;
                    }

                    if (candidate.Reasons.Count > best.Reasons.Count)
                    {
                        best = candidate;
                        continue;
                    }

                    if (candidate.Record.LastUpdatedOn > best.Record.LastUpdatedOn)
                    {
                        best = candidate;
                    }
                }
            }

            if (best is null)
            {
                diagnostics.Add("no candidate scored above zero");
                return null;
            }

            if (best.Score < MinimumCandidateScore)
            {
                diagnostics.Add($"best score {best.Score:F1} below threshold {MinimumCandidateScore:F1} for event '{best.Record.EventId}'");
                return null;
            }

            DebugLogger.Log($"SelectEventCandidate chose Event='{best.Record.EventId}' score={best.Score:F1} reasons={best.ReasonSummary} details={best.DetailSummary}");

            return best;
        }

        private static string SummarizeDiagnostics(IReadOnlyList<string> entries)
        {
            if (entries is null || entries.Count == 0)
            {
                return "(none)";
            }

            const int MaxItems = 3;
            var sample = entries.Take(MaxItems).ToList();
            var suffix = entries.Count > MaxItems ? " | ..." : string.Empty;
            return string.Join(" | ", sample) + suffix;
        }

        private static string NormalizeSubjectForFallback(string? subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
            {
                return string.Empty;
            }

            var value = subject!.Trim();
            while (value.StartsWith("RE:", StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("FW:", StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("FWD:", StringComparison.OrdinalIgnoreCase))
            {
                var separatorIndex = value.IndexOf(':');
                if (separatorIndex < 0)
                {
                    break;
                }

                value = value.Substring(separatorIndex + 1).TrimStart();
            }

            if (value.Length == 0)
            {
                return string.Empty;
            }

            var parts = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length == 0 ? string.Empty : string.Join(" ", parts).ToUpperInvariant();
        }

        private static bool IsSameMail(EmailItem existing, EmailItem candidate)
        {
            if (!string.IsNullOrWhiteSpace(candidate.EntryId) && string.Equals(existing.EntryId, candidate.EntryId, StringComparison.OrdinalIgnoreCase))
            {
                DebugLogger.Log($"IsSameMail matched by EntryId existing='{existing.EntryId}' candidate='{candidate.EntryId}'");
                return true;
            }

            var existingMessageId = NormalizeMessageId(existing.InternetMessageId);
            var candidateMessageId = NormalizeMessageId(candidate.InternetMessageId);
            if (!string.IsNullOrWhiteSpace(candidateMessageId) && !string.IsNullOrWhiteSpace(existingMessageId) &&
                string.Equals(existingMessageId, candidateMessageId, StringComparison.OrdinalIgnoreCase))
            {
                DebugLogger.Log($"IsSameMail matched by InternetMessageId existing='{existing.InternetMessageId}' candidate='{candidate.InternetMessageId}'");
                return true;
            }

            if (!string.IsNullOrWhiteSpace(candidate.ConversationId) && string.Equals(existing.ConversationId, candidate.ConversationId, StringComparison.OrdinalIgnoreCase))
            {
                var missingEntryIds = string.IsNullOrWhiteSpace(existing.EntryId) && string.IsNullOrWhiteSpace(candidate.EntryId);
                var missingMessageIds = string.IsNullOrWhiteSpace(existing.InternetMessageId) && string.IsNullOrWhiteSpace(candidate.InternetMessageId);
                if (!(missingEntryIds && missingMessageIds))
                {
                    return false;
                }

                var sameSender = string.Equals(existing.Sender, candidate.Sender, StringComparison.OrdinalIgnoreCase);
                var sameSubject = string.Equals(existing.Subject, candidate.Subject, StringComparison.OrdinalIgnoreCase);
                var withinWindow = Math.Abs((existing.ReceivedOn - candidate.ReceivedOn).TotalMinutes) <= DeduplicationWindow.TotalMinutes;
                if (sameSender && sameSubject && withinWindow)
                {
                    DebugLogger.Log($"IsSameMail matched by fallback Conversation='{candidate.ConversationId}' senderMatch={sameSender} subjectMatch={sameSubject} deltaSeconds={Math.Abs((existing.ReceivedOn - candidate.ReceivedOn).TotalSeconds)}");
                    return true;
                }
            }

            var existingRoot = GetThreadRoot(existing.ThreadIndex);
            var candidateRoot = GetThreadRoot(candidate.ThreadIndex);
            if (!string.IsNullOrEmpty(existingRoot) && string.Equals(existingRoot, candidateRoot, StringComparison.Ordinal))
            {
                var candidateFingerprint = candidate.BodyFingerprint ?? string.Empty;
                if (!string.IsNullOrEmpty(candidateFingerprint) && MailBodyFingerprint.IsSimilar(candidateFingerprint, new[] { existing.BodyFingerprint }))
                {
                    DebugLogger.Log($"IsSameMail matched by ThreadIndex root '{existingRoot}' using body fingerprint similarity");
                    return true;
                }
            }

            return false;
        }

        private static string GetThreadRoot(string? threadIndex)
        {
            if (string.IsNullOrWhiteSpace(threadIndex))
            {
                return string.Empty;
            }

            const int RootLength = 44;
            var normalized = threadIndex!.Trim();
            return normalized.Length >= RootLength ? normalized.Substring(0, RootLength) : normalized;
        }

        private static string GetThreadIndexHint(string? threadIndexPrefix, string? threadIndex)
        {
            var candidate = !string.IsNullOrWhiteSpace(threadIndexPrefix) ? threadIndexPrefix : threadIndex;
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return string.Empty;
            }

            var normalized = candidate!.Trim();
            return normalized.Length <= 4 ? normalized : normalized.Substring(0, 4);
        }

        private static string NormalizeMessageId(string? value)
        {
            if (value is null)
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            if (trimmed.Length >= 2 && trimmed[0] == '<' && trimmed[trimmed.Length - 1] == '>')
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2).Trim();
            }

            return trimmed;
        }

        private static string[] NormalizeMessageIds(IEnumerable<string> values)
        {
            var normalized = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in values ?? Array.Empty<string>())
            {
                var candidate = NormalizeMessageId(value);
                if (!string.IsNullOrEmpty(candidate))
                {
                    normalized.Add(candidate);
                }
            }

            return normalized.Count == 0 ? Array.Empty<string>() : normalized.ToArray();
        }

        private static bool EnsureConversationTracked(EventRecord record, string? conversationId)
        {
            if (record is null)
            {
                return false;
            }

            var normalized = conversationId?.Trim();
            if (string.IsNullOrEmpty(normalized))
            {
                return false;
            }

            if (record.ConversationIds.Any(id => string.Equals(id, normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            record.ConversationIds.Add(normalized!);
            return true;
        }

        private static void AppendAttachments(EventRecord record, IEnumerable<AttachmentItem> attachments)
        {
            foreach (var attachment in attachments)
            {
                if (record.Attachments.Any(a => string.Equals(a.Id, attachment.Id, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                record.Attachments.Add(attachment);
            }
        }

        private void LoadFromDisk()
        {
            if (!File.Exists(_storePath))
            {
                return;
            }

            var json = File.ReadAllText(_storePath);
            var records = JsonConvert.DeserializeObject<List<EventRecord>>(json, _serializerSettings);
            if (records is null)
            {
                return;
            }

            _events.Clear();
            _events.AddRange(records);
        }

        private async Task PersistLockedAsync()
        {
            await Task.Run(() =>
            {
                var json = JsonConvert.SerializeObject(_events, _serializerSettings);
                File.WriteAllText(_storePath, json);
            }).ConfigureAwait(false);
        }

        private static EventRecord Clone(EventRecord record)
        {
            var json = JsonConvert.SerializeObject(record);
            return JsonConvert.DeserializeObject<EventRecord>(json)!;
        }

        private static void CopyInto(EventRecord source, EventRecord target)
        {
            target.EventTitle = source.EventTitle;
            target.DashboardTemplateId = source.DashboardTemplateId;
            target.Status = source.Status;
            target.PriorityLevel = source.PriorityLevel;
            target.DisplayColumnSource = source.DisplayColumnSource;
            target.DisplayColumnCustomValue = source.DisplayColumnCustomValue;
            target.ConversationIds = source.ConversationIds.ToList();
            target.DashboardItems = source.DashboardItems.Select(item => new DashboardItem { Key = item.Key, Value = item.Value }).ToList();
            target.Emails = source.Emails.Select(email => new EmailItem
            {
                EntryId = email.EntryId,
                StoreId = email.StoreId,
                ConversationId = email.ConversationId,
                InternetMessageId = email.InternetMessageId,
                Sender = email.Sender,
                To = email.To,
                Subject = email.Subject,
                Participants = email.Participants ?? Array.Empty<string>(),
                BodyFingerprint = email.BodyFingerprint,
                ThreadIndex = email.ThreadIndex,
                ThreadIndexPrefix = email.ThreadIndexPrefix,
                ReferenceMessageIds = email.ReferenceMessageIds ?? Array.Empty<string>(),
                ReceivedOn = email.ReceivedOn,
                IsNewOrUpdated = email.IsNewOrUpdated,
                IsRemoved = email.IsRemoved
            }).ToList();
            target.Attachments = source.Attachments.Select(attachment => new AttachmentItem
            {
                Id = attachment.Id,
                FileName = attachment.FileName,
                FileType = attachment.FileType,
                FileSizeBytes = attachment.FileSizeBytes,
                SourceMailEntryId = attachment.SourceMailEntryId
            }).ToList();
            target.AdditionalFiles = new List<string>(source.AdditionalFiles ?? Enumerable.Empty<string>());
            target.RelatedSubjects = new HashSet<string>(source.RelatedSubjects ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            target.Participants = new HashSet<string>(source.Participants ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            target.ProcessedMessageIds = new HashSet<string>(source.ProcessedMessageIds ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            target.LastUpdatedOn = source.LastUpdatedOn;
        }

        private void RaiseChanged(EventRecord record, string reason)
        {
            DebugLogger.Log($"EventRepository.RaiseChanged: Triggering update for Event='{record.EventId}' Reason='{reason}'");
            var handlers = EventChanged;
            if (handlers is null)
            {
                DebugLogger.Log($"EventRepository.RaiseChanged: No subscribers for Event='{record.EventId}'");
                return;
            }

            var args = new EventChangedEventArgs(record, reason);

            void RaiseEvent(object? _)
            {
                handlers(this, args);
            }

            if (_syncContext is null || _syncContext == SynchronizationContext.Current)
            {
                RaiseEvent(null);
            }
            else
            {
                _syncContext.Post(RaiseEvent, null);
            }
        }

        private static string BuildDefaultStorePath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(root, "OSEM", "event-store.json");
        }

        private sealed class EventMatchCandidate
        {
            private readonly HashSet<string> _reasonSet = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _detailSet = new(StringComparer.Ordinal);

            public EventMatchCandidate(EventRecord record)
            {
                Record = record ?? throw new ArgumentNullException(nameof(record));
            }

            public EventRecord Record { get; }
            public double Score { get; private set; }
            public List<string> Reasons { get; } = new();
            public List<string> Details { get; } = new();
            public bool PreferredBiasApplied { get; private set; }
            public bool HasNewContent { get; set; }

            public string ReasonSummary => Reasons.Count == 0 ? "(none)" : string.Join(", ", Reasons);
            public string DetailSummary => Details.Count == 0 ? "(none)" : string.Join(" | ", Details);

            public void AddScore(double weight, string reason, string detail)
            {
                Score += weight;

                if (!string.IsNullOrEmpty(reason) && _reasonSet.Add(reason))
                {
                    Reasons.Add(reason);
                }

                if (!string.IsNullOrEmpty(detail) && _detailSet.Add(detail))
                {
                    Details.Add(detail);
                }
            }

            public void ApplyPreferredBias(double weight, string detail)
            {
                Score += weight;
                PreferredBiasApplied = true;

                if (_reasonSet.Add("preferred"))
                {
                    Reasons.Add("preferred");
                }

                if (!string.IsNullOrEmpty(detail) && _detailSet.Add(detail))
                {
                    Details.Add(detail);
                }
            }
        }
    }

    internal static class EventRecordExtensions
    {
        public static TResult? Let<TSource, TResult>(this TSource? source, Func<TSource, TResult> map) where TSource : class where TResult : class
        {
            return source is null ? null : map(source);
        }
    }
}
