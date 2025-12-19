using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OSEMAddIn.ViewModels
{
    public class FileTypeOption : ViewModelBase
    {
        private bool _isSelected;

        public string Label { get; set; } = string.Empty;
        public HashSet<string> Extensions { get; set; } = new HashSet<string>();

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
