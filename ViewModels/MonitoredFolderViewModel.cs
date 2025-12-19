namespace OSEMAddIn.ViewModels
{
    public class MonitoredFolderViewModel : ViewModelBase
    {
        private string _name = string.Empty;
        private string _entryId = string.Empty;
        private string _path = string.Empty;

        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        public string EntryId
        {
            get => _entryId;
            set => SetProperty(ref _entryId, value);
        }

        public string Path
        {
            get => _path;
            set => SetProperty(ref _path, value);
        }
    }
}
