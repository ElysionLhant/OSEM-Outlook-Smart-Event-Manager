using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    internal class TemplateRuleViewModel : ViewModelBase
    {
        private string _participant = string.Empty;
        private DashboardTemplate _selectedTemplate = new DashboardTemplate();

        public string Participant
        {
            get => _participant;
            set { _participant = value; RaisePropertyChanged(); }
        }

        public DashboardTemplate SelectedTemplate
        {
            get => _selectedTemplate;
            set { _selectedTemplate = value; RaisePropertyChanged(); }
        }
    }
}
