using System.IO;

namespace OSEMAddIn.ViewModels
{
    public class TemplateFileViewModel : ViewModelBase
    {
        private bool _isSelected;
        private string _filePath;

        public TemplateFileViewModel(string filePath)
        {
            _filePath = filePath;
        }

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath == value) return;
                _filePath = value;
                RaisePropertyChanged();
                RaisePropertyChanged(nameof(FileName));
            }
        }

        public string FileName => Path.GetFileName(FilePath);

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                RaisePropertyChanged();
            }
        }

        public string OriginalPath { get; set; } = string.Empty;
        public bool IsCommonFile { get; set; }
    }
}
