using System;
using Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn.Services
{
    internal sealed class ServiceContainer : IDisposable
    {
        public ServiceContainer(Application outlookApplication)
        {
            OutlookApplication = outlookApplication ?? throw new ArgumentNullException(nameof(outlookApplication));
            EventRepository = new EventRepository(OutlookApplication);
            DashboardTemplates = new DashboardTemplateService();
            PromptLibrary = new PromptLibraryService();
            PythonScripts = new PythonScriptService();
            RegexExtraction = new RegexExtractionService();
            LlmExtraction = new LlmExtractionService();
            LlmConfigurations = new LlmConfigurationService();
            OllamaModels = new OllamaModelService();
            CsvExport = new CsvExportService();
            EmailTemplates = new EmailTemplateService();
            TemplatePreferences = new TemplatePreferenceService();
            TemplatePackages = new TemplatePackageService(DashboardTemplates, PromptLibrary, PythonScripts);
            BackupService = new BackupService(EventRepository, DashboardTemplates, PromptLibrary, PythonScripts, EmailTemplates);
            EventMonitor = new OutlookEventMonitor(OutlookApplication, EventRepository);
            BusyState = new BusyStateService();
        }

        public Application OutlookApplication { get; }
        public EventRepository EventRepository { get; }
        public DashboardTemplateService DashboardTemplates { get; }
        public PromptLibraryService PromptLibrary { get; }
        public PythonScriptService PythonScripts { get; }
        public TemplatePackageService TemplatePackages { get; }
        public BackupService BackupService { get; }
        public RegexExtractionService RegexExtraction { get; }
        public LlmExtractionService LlmExtraction { get; }
        public LlmConfigurationService LlmConfigurations { get; }
        public OllamaModelService OllamaModels { get; }
        public CsvExportService CsvExport { get; }
        public EmailTemplateService EmailTemplates { get; }
        public TemplatePreferenceService TemplatePreferences { get; }
        public OutlookEventMonitor EventMonitor { get; }
        public BusyStateService BusyState { get; }

        public void Dispose()
        {
            EventMonitor.Dispose();
            EventRepository.Dispose();
        }
    }
}
