using System;
using OSEMAddIn.Models;

namespace OSEMAddIn.Services
{
    internal sealed class EventChangedEventArgs : EventArgs
    {
        public EventChangedEventArgs(EventRecord record, string changeReason)
        {
            Record = record;
            ChangeReason = changeReason;
        }

        public EventRecord Record { get; }
        public string ChangeReason { get; }
    }
}
