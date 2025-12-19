using System.Windows;
using System.Windows.Controls;
using OSEMAddIn.ViewModels;

namespace OSEMAddIn.Views.EventManager
{
    public partial class TemplateEditorView : Window
    {
        public TemplateEditorView()
        {
            InitializeComponent();
        }

        private void MoreOptions_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void TemplateList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is TemplateEditorViewModel vm)
            {
                vm.IsMultiSelection = TemplateList.SelectedItems.Count > 1;
            }
        }

        private void TemplateFilesList_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (DataContext is TemplateEditorViewModel vm)
                {
                    vm.AddFiles(files);
                }
            }
        }

        private void InsertVariable_Click(object sender, RoutedEventArgs e)
        {
            if (VariableList.SelectedItem is PromptVariable variable && PromptBodyBox != null)
            {
                int caretIndex = PromptBodyBox.CaretIndex;
                string textToInsert = variable.InsertText;
                
                // Insert text
                string currentText = PromptBodyBox.Text ?? string.Empty;
                if (caretIndex < 0) caretIndex = 0;
                if (caretIndex > currentText.Length) caretIndex = currentText.Length;

                PromptBodyBox.Text = currentText.Insert(caretIndex, textToInsert);
                
                // Restore focus and move caret
                PromptBodyBox.Focus();
                PromptBodyBox.CaretIndex = caretIndex + textToInsert.Length;
                
                // Force binding update because we modified Text property programmatically
                var binding = PromptBodyBox.GetBindingExpression(TextBox.TextProperty);
                binding?.UpdateSource();
            }
        }

        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox comboBox)
            {
                // Use dispatcher to ensure this happens after focus is fully set
                Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    var textBox = (TextBox)comboBox.Template.FindName("PART_EditableTextBox", comboBox);
                    if (textBox != null)
                    {
                        textBox.SelectAll();
                    }
                }));
            }
        }

        private void ScriptConfig_LostFocus(object sender, RoutedEventArgs e)
        {
            if (DataContext is TemplateEditorViewModel vm)
            {
                vm.SaveScript();
            }
        }
    }
}
