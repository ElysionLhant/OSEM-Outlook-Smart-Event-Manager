using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal interface IEventRepository
    {
        event EventHandler<EventChangedEventArgs>? EventChanged;

        Task<IReadOnlyList<EventRecord>> GetAllAsync();
        Task<EventRecord?> GetByIdAsync(string eventId);
        EventRecord? GetEvent(string eventId);
        Task<EventRecord> CreateFromMailAsync(MailItem mailItem, string? dashboardTemplateId = null, IEnumerable<string>? knownParticipants = null);
        Task UpdateAsync(EventRecord record);
        Task ImportAsync(EventRecord record);
        Task ArchiveAsync(IEnumerable<string> eventIds);
        Task ReopenAsync(string eventId);
        Task DeleteAsync(string eventId);
        Task DeleteAsync(IEnumerable<string> eventIds);
        Task MarkMessageIdsAsNotFoundAsync(string eventId, IEnumerable<string> messageIds);
        Task<EventRecord?> TryAddMailAsync(MailSnapshot snapshot, string? preferredEventId = null);
    }
}
