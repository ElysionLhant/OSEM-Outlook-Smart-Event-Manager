using System.Windows;

namespace OSEMAddIn.Views.EventManager
{
    public partial class TemplateRuleManagerView : Window
    {
        public TemplateRuleManagerView()
        {
            InitializeComponent();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
