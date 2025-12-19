using System;
using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    internal sealed class MailItemViewModel : ViewModelBase
    {
        private bool _isNewOrUpdated;

        public MailItemViewModel(EmailItem email)
        {
            EntryId = email.EntryId;
            StoreId = email.StoreId;
            ConversationId = email.ConversationId;
            InternetMessageId = email.InternetMessageId;
            Sender = email.Sender;
            To = email.To;
            Subject = email.Subject;
            ReceivedOn = email.ReceivedOn.ToLocalTime();
            _isNewOrUpdated = email.IsNewOrUpdated;
        }

        public string EntryId { get; }
        public string StoreId { get; }
        public string ConversationId { get; }
        public string InternetMessageId { get; }
        public string Sender { get; }
        public string To { get; }
        public string Subject { get; }
        public DateTime ReceivedOn { get; }

        public bool IsNewOrUpdated
        {
            get => _isNewOrUpdated;
            set
            {
                if (_isNewOrUpdated == value)
                {
                    return;
                }

                _isNewOrUpdated = value;
                RaisePropertyChanged();
            }
        }
    }
}
