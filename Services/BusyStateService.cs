using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace OSEMAddIn.Services
{
    public class BusyStateService : INotifyPropertyChanged
    {
        private bool _isBusy;
        private string _busyMessage = string.Empty;

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged();
                }
            }
        }

        public string BusyMessage
        {
            get => _busyMessage;
            private set
            {
                if (_busyMessage != value)
                {
                    _busyMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public void SetBusy(string message)
        {
            BusyMessage = message;
            IsBusy = true;
        }

        public void ClearBusy()
        {
            IsBusy = false;
            BusyMessage = string.Empty;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
