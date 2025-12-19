using System.Collections.ObjectModel;
using System.Linq;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;
using OSEMAddIn.Services;

namespace OSEMAddIn.ViewModels
{
    internal class TemplateRuleManagerViewModel : ViewModelBase
    {
        private readonly ServiceContainer _services;
        private TemplateRuleViewModel? _selectedRule;

        public TemplateRuleManagerViewModel(ServiceContainer services)
        {
            _services = services;
            Rules = new ObservableCollection<TemplateRuleViewModel>();
            Templates = new ObservableCollection<DashboardTemplate>(_services.DashboardTemplates.GetTemplates());
            
            LoadRules();

            AddRuleCommand = new RelayCommand(_ => AddRule());
            RemoveRuleCommand = new RelayCommand(_ => RemoveRule(), _ => SelectedRule != null);
            SaveCommand = new RelayCommand(_ => Save());
        }

        public ObservableCollection<TemplateRuleViewModel> Rules { get; }
        public ObservableCollection<DashboardTemplate> Templates { get; }

        public TemplateRuleViewModel? SelectedRule
        {
            get => _selectedRule;
            set
            {
                _selectedRule = value;
                RaisePropertyChanged();
                RemoveRuleCommand.RaiseCanExecuteChanged();
            }
        }

        public RelayCommand AddRuleCommand { get; }
        public RelayCommand RemoveRuleCommand { get; }
        public RelayCommand SaveCommand { get; }

        private void LoadRules()
        {
            Rules.Clear();
            var prefs = _services.TemplatePreferences.GetPreferences();
            
            // Group by TemplateId to allow multi-participant rules
            var grouped = prefs.GroupBy(p => p.Value);

            foreach (var group in grouped)
            {
                var template = _services.DashboardTemplates.FindById(group.Key);
                if (template != null)
                {
                    // Join participants with semicolon
                    var participants = string.Join("; ", group.Select(p => p.Key));
                    Rules.Add(new TemplateRuleViewModel
                    {
                        Participant = participants,
                        SelectedTemplate = template
                    });
                }
            }
        }

        private void AddRule()
        {
            Rules.Add(new TemplateRuleViewModel
            {
                Participant = "new@example.com; other@example.com",
                SelectedTemplate = Templates.FirstOrDefault()
            });
        }

        private void RemoveRule()
        {
            if (SelectedRule != null)
            {
                Rules.Remove(SelectedRule);
            }
        }

        private void Save()
        {
            _services.TemplatePreferences.Clear();
            foreach (var rule in Rules)
            {
                if (!string.IsNullOrWhiteSpace(rule.Participant) && rule.SelectedTemplate != null)
                {
                    // Split by common separators
                    var participants = rule.Participant.Split(new[] { ';', ',', '，', '；' }, System.StringSplitOptions.RemoveEmptyEntries);
                    
                    foreach (var p in participants)
                    {
                        var clean = MailParticipantExtractor.Normalize(p);
                        if (!string.IsNullOrEmpty(clean))
                        {
                            _services.TemplatePreferences.UpdatePreference(clean, rule.SelectedTemplate.TemplateId);
                        }
                    }
                }
            }
        }
    }
}
