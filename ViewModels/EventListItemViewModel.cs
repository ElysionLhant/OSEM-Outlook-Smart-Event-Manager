using System;
using System.Windows.Input;
using System.Windows.Media;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;
using OSEMAddIn.Services;

namespace OSEMAddIn.ViewModels
{
    internal sealed class EventListItemViewModel : ViewModelBase
    {
        private readonly IEventRepository _repository;
        private readonly EventRecord _record;
        private bool _isSelected;
        private bool _hasNewMail;
        private int _priorityLevel;

        public EventListItemViewModel(EventRecord record, IEventRepository repository)
        {
            _record = record ?? throw new ArgumentNullException(nameof(record));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));

            EventId = record.EventId;
            EventTitle = record.EventTitle;
            Status = record.Status;
            LastUpdatedOn = record.LastUpdatedOn.ToLocalTime();
            DashboardTemplateId = record.DashboardTemplateId;
            PriorityLevel = record.PriorityLevel;

            TogglePriorityCommand = new RelayCommand(_ => TogglePriority());

            UpdateDisplayColumn(record);
            RefreshIndicators(record);
        }

        public string EventId { get; }
        public string EventTitle { get; private set; }
        public string DisplayColumnText { get; private set; } = string.Empty;
        public EventStatus Status { get; private set; }
        public DateTime LastUpdatedOn { get; private set; }
        public string DashboardTemplateId { get; private set; }

        public ICommand TogglePriorityCommand { get; }

        public int PriorityLevel
        {
            get => _priorityLevel;
            private set
            {
                if (_priorityLevel == value) return;
                _priorityLevel = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(PriorityColor));
                RaisePropertyChanged(nameof(PriorityIcon));
                RaisePropertyChanged(nameof(HasPriority));
            }
        }

        public string PriorityIcon => _priorityLevel == 0 ? "☆" : "★";

        public bool HasPriority => _priorityLevel > 0;

        public Brush PriorityColor
        {
            get
            {
                return _priorityLevel switch
                {
                    1 => Brushes.Yellow,
                    2 => Brushes.Orange,
                    3 => Brushes.Red,
                    _ => Brushes.Transparent
                };
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                RaisePropertyChanged();
            }
        }

        public bool HasNewMail
        {
            get => _hasNewMail;
            private set
            {
                if (_hasNewMail == value)
                {
                    return;
                }

                _hasNewMail = value;
                RaisePropertyChanged();
            }
        }

        public void Update(EventRecord record)
        {
            EventTitle = record.EventTitle;
            Status = record.Status;
            LastUpdatedOn = record.LastUpdatedOn.ToLocalTime();
            DashboardTemplateId = record.DashboardTemplateId;
            PriorityLevel = record.PriorityLevel;
            UpdateDisplayColumn(record);
            RefreshIndicators(record);
            RaisePropertyChanged(nameof(EventTitle));
            RaisePropertyChanged(nameof(DisplayColumnText));
            RaisePropertyChanged(nameof(Status));
            RaisePropertyChanged(nameof(LastUpdatedOn));
        }

        public void MarkMailAsRead()
        {
            HasNewMail = false;
        }

        private async void TogglePriority()
        {
            _record.PriorityLevel = (_record.PriorityLevel + 1) % 4;
            PriorityLevel = _record.PriorityLevel;
            await _repository.UpdateAsync(_record);
        }

        private void UpdateDisplayColumn(EventRecord record)
        {
            var item = record.DashboardItems.Find(i => string.Equals(i.Key, record.DisplayColumnSource, StringComparison.OrdinalIgnoreCase));
            if (item != null)
            {
                DisplayColumnText = $"{item.Key}: {item.Value}";
            }
            else if (string.Equals(record.DisplayColumnSource, "Custom", StringComparison.OrdinalIgnoreCase))
            {
                DisplayColumnText = record.DisplayColumnCustomValue;
            }
            else
            {
                DisplayColumnText = record.DisplayColumnSource;
            }
        }

        private void RefreshIndicators(EventRecord record)
        {
            HasNewMail = record.Emails.Exists(mail => mail.IsNewOrUpdated);
        }
    }
}
