#nullable enable
using System;
using System.Threading.Tasks;
using System.Windows.Input;
using OSEMAddIn.Commands;
using OSEMAddIn.Services;

namespace OSEMAddIn.ViewModels
{
    internal sealed class ShellViewModel : ViewModelBase
    {
        private readonly ServiceContainer _services;
        private ViewModelBase _currentViewModel;
        private bool _isForcefullyActive;

        public ShellViewModel(ServiceContainer services)
        {
            _services = services ?? throw new ArgumentNullException(nameof(services));
            EventManager = new EventManagerViewModel(_services);
            EventDetail = new EventDetailViewModel(_services);
            ForceActivateCommand = new RelayCommand(_ => ForceActivate());

            EventManager.OpenEventRequested += OnOpenEventRequested;
            EventDetail.BackRequested += OnBackRequested;

            _services.BusyState.PropertyChanged += OnBusyStateChanged;

            _currentViewModel = EventManager;
            _ = InitializeAsync();
        }

        public EventManagerViewModel EventManager { get; }
        public EventDetailViewModel EventDetail { get; }
        public ICommand ForceActivateCommand { get; }

        public bool IsBusy => _services.BusyState.IsBusy && !_isForcefullyActive;
        public string BusyMessage => _services.BusyState.BusyMessage;

        public ViewModelBase CurrentViewModel
        {
            get => _currentViewModel;
            private set
            {
                if (_currentViewModel == value)
                {
                    return;
                }

                _currentViewModel = value;
                RaisePropertyChanged();
            }
        }

        private void OnBusyStateChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(BusyStateService.IsBusy))
            {
                if (!_services.BusyState.IsBusy)
                {
                    _isForcefullyActive = false;
                }
                RaisePropertyChanged(nameof(IsBusy));
            }
            else if (e.PropertyName == nameof(BusyStateService.BusyMessage))
            {
                RaisePropertyChanged(nameof(BusyMessage));
            }
        }

        private void ForceActivate()
        {
            _isForcefullyActive = true;
            RaisePropertyChanged(nameof(IsBusy));
        }

        private async Task InitializeAsync()
        {
            _services.EventMonitor.SearchStatusChanged += (s, msg) =>
            {
                if (msg == "Searching...")
                {
                    _services.BusyState.SetBusy("Searching for missing emails...");
                }
                else if (msg == "Waiting for Outlook sync...")
                {
                    _services.BusyState.SetBusy("Waiting for Outlook sync...");
                }
                else
                {
                    _services.BusyState.ClearBusy();
                }
            };

            _services.EventMonitor.Start();
            await EventManager.InitializeAsync();
        }

        private async void OnOpenEventRequested(object? sender, string eventId)
        {
            DebugLogger.Log($"ShellViewModel: OpenEventRequested for ID: {eventId}");
            CurrentViewModel = EventDetail;
            await EventDetail.LoadAsync(eventId);
        }

        private void OnBackRequested(object? sender, string _)
        {
            CurrentViewModel = EventManager;
        }
    }
}
