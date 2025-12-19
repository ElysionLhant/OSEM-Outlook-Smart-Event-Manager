using System.Windows;

namespace OSEMAddIn.Views
{
    public partial class ConflictResolutionDialog : Window
    {
        public enum Resolution { Overwrite, Skip, Rename }
        public Resolution Result { get; private set; } = Resolution.Skip;

        public string ItemName
        {
            get { return (string)GetValue(ItemNameProperty); }
            set { SetValue(ItemNameProperty, value); }
        }

        public static readonly DependencyProperty ItemNameProperty =
            DependencyProperty.Register("ItemName", typeof(string), typeof(ConflictResolutionDialog), new PropertyMetadata(string.Empty));

        public ConflictResolutionDialog(string itemName)
        {
            InitializeComponent();
            ItemName = itemName;
            DataContext = this;
        }

        private void Overwrite_Click(object sender, RoutedEventArgs e)
        {
            Result = Resolution.Overwrite;
            DialogResult = true;
        }

        private void Skip_Click(object sender, RoutedEventArgs e)
        {
            Result = Resolution.Skip;
            DialogResult = true;
        }

        private void Rename_Click(object sender, RoutedEventArgs e)
        {
            Result = Resolution.Rename;
            DialogResult = true;
        }
    }
}
