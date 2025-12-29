#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using OSEMAddIn.Commands;
using OSEMAddIn.Models;

namespace OSEMAddIn.ViewModels
{
    public class ExportOptionsViewModel : ViewModelBase
    {
        private DashboardTemplate? _selectedTemplate;
        private DateTime _startDate;
        private DateTime _endDate;
        private bool _exportDashboardData = true;
        private bool _exportEventAttachments = true;
        private bool _exportAdditionalFiles = false;
        private bool _exportAllFileTypes = false;
        private string _targetPath = string.Empty;
        private string _folderNamingMode = "Event Title";
        private string _selectedNamingKey = string.Empty;
        private bool _exportInProgress = false;
        private bool _exportArchived = true;

        public ExportOptionsViewModel(IEnumerable<DashboardTemplate> templates, DashboardTemplate? initialTemplate, DateTime startDate, DateTime endDate)
        {
            NamingOptions = new ObservableCollection<string> { "Event Title" };
            Templates = new ObservableCollection<DashboardTemplate>(templates);
            SelectedTemplate = initialTemplate ?? Templates.FirstOrDefault();
            StartDate = startDate;
            EndDate = endDate;

            InitializeFileTypes();

            BrowseCommand = new RelayCommand(_ => BrowseFolder());
            ExportCommand = new RelayCommand(_ => RequestClose?.Invoke(true), _ => CanExport());
            
            UpdateNamingOptions();
        }

        private void InitializeFileTypes()
        {
            FileTypeOptions = new ObservableCollection<FileTypeOption>
            {
                new FileTypeOption 
                { 
                    Label = Properties.Resources.PDF_Document_pdf, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "pdf" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.Word_Document_doc_docx, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "doc", "docx" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.Excel_Spreadsheet_xls_xlsx_csv, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "xls", "xlsx", "csv" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.PPT_Presentation_ppt_pptx, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "ppt", "pptx" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.Image_png_jpg_bmp, 
                    IsSelected = false, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "png", "jpg", "jpeg", "bmp", "gif", "tif", "tiff" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.Archive_zip_rar_7z, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "zip", "rar", "7z", "tar", "gz" } 
                },
                new FileTypeOption 
                { 
                    Label = Properties.Resources.Text_File_txt_md_rtf, 
                    IsSelected = true, 
                    Extensions = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "txt", "md", "rtf", "log" } 
                }
            };
        }

        public event Action<bool>? RequestClose;
        public event Action? BrowseRequested;

        public ObservableCollection<DashboardTemplate> Templates { get; }
        public ObservableCollection<FileTypeOption> FileTypeOptions { get; private set; } = new();

        public DashboardTemplate? SelectedTemplate
        {
            get => _selectedTemplate;
            set
            {
                if (SetProperty(ref _selectedTemplate, value))
                {
                    UpdateNamingOptions();
                }
            }
        }

        public DateTime StartDate
        {
            get => _startDate;
            set => SetProperty(ref _startDate, value);
        }

        public DateTime EndDate
        {
            get => _endDate;
            set => SetProperty(ref _endDate, value);
        }

        public bool ExportDashboardData
        {
            get => _exportDashboardData;
            set => SetProperty(ref _exportDashboardData, value);
        }

        public bool ExportEventAttachments
        {
            get => _exportEventAttachments;
            set => SetProperty(ref _exportEventAttachments, value);
        }

        public bool ExportAdditionalFiles
        {
            get => _exportAdditionalFiles;
            set => SetProperty(ref _exportAdditionalFiles, value);
        }

        public bool ExportAllFileTypes
        {
            get => _exportAllFileTypes;
            set
            {
                if (SetProperty(ref _exportAllFileTypes, value))
                {
                    RaisePropertyChanged(nameof(AreSpecificFileTypesEnabled));
                }
            }
        }

        public bool ExportInProgress
        {
            get => _exportInProgress;
            set => SetProperty(ref _exportInProgress, value);
        }

        public bool ExportArchived
        {
            get => _exportArchived;
            set => SetProperty(ref _exportArchived, value);
        }

        public bool AreSpecificFileTypesEnabled => !ExportAllFileTypes;

        public string TargetPath
        {
            get => _targetPath;
            set
            {
                if (SetProperty(ref _targetPath, value))
                {
                    ExportCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public ObservableCollection<string> NamingOptions { get; }

        public string FolderNamingMode
        {
            get => _folderNamingMode;
            set => SetProperty(ref _folderNamingMode, value);
        }

        public ICommand BrowseCommand { get; }
        public RelayCommand ExportCommand { get; }

        private void BrowseFolder()
        {
            BrowseRequested?.Invoke();
        }

        private bool CanExport()
        {
            return !string.IsNullOrWhiteSpace(TargetPath) && SelectedTemplate != null;
        }

        private void UpdateNamingOptions()
        {
            var currentSelection = FolderNamingMode;
            NamingOptions.Clear();
            NamingOptions.Add("Event Title");

            if (SelectedTemplate != null)
            {
                if (SelectedTemplate.Fields != null)
                {
                     foreach (var field in SelectedTemplate.Fields)
                     {
                         NamingOptions.Add($"Key: {field}");
                     }
                }
            }

            if (NamingOptions.Contains(currentSelection))
            {
                FolderNamingMode = currentSelection;
            }
            else
            {
                FolderNamingMode = NamingOptions.First();
            }
        }
    }
}
