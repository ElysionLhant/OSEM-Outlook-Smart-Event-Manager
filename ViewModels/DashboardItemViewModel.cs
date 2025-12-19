using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    internal sealed class DashboardItemViewModel : ViewModelBase
    {
        private string _value;
        private readonly System.Action? _onChanged;

        public DashboardItemViewModel(string key, string value, System.Action? onChanged = null)
        {
            Key = key;
            _value = value;
            _onChanged = onChanged;
        }

        public string Key { get; }

        public string Value
        {
            get => _value;
            set
            {
                if (_value == value)
                {
                    return;
                }

                _value = value;
                RaisePropertyChanged();
                _onChanged?.Invoke();
            }
        }

        public DashboardItem ToModel() => new() { Key = Key, Value = Value };
    }
}
