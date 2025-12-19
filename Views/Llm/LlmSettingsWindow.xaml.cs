using System;
using System.Threading.Tasks;
using System.Windows;
using OSEMAddIn.ViewModels;

namespace OSEMAddIn.Views.Llm
{
    public partial class LlmSettingsWindow : Window
    {
        private readonly LlmSettingsViewModel _viewModel;

        internal LlmSettingsWindow(LlmSettingsViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;
            Loaded += OnLoaded;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= OnLoaded;
            await InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                await _viewModel.InitializeAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(Properties.Resources.Failed_to_load_Ollama_model_list_ex_Message, ex.Message), "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnSaveClicked(object sender, RoutedEventArgs e)
        {
            _viewModel.Save();
            DialogResult = true;
            Close();
        }

        private void OnClearTemplateClicked(object sender, RoutedEventArgs e)
        {
            _viewModel.ClearTemplateOverride();
        }

        private void OnCloseClicked(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
