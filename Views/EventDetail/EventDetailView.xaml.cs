#nullable enable
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.Outlook;
using OSEMAddIn.ViewModels;
using OSEMAddIn.Views.Llm;

namespace OSEMAddIn.Views.EventDetail
{
    public partial class EventDetailView : UserControl
    {
        public EventDetailView()
        {
            InitializeComponent();
            DataContextChanged += OnDataContextChanged;
        }

        private EventDetailViewModel? ViewModel => DataContext as EventDetailViewModel;

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.OldValue is EventDetailViewModel oldVm)
            {
                oldVm.ManagePromptRequested -= OnManagePromptRequested;
            }

            if (e.NewValue is EventDetailViewModel newVm)
            {
                newVm.ManagePromptRequested += OnManagePromptRequested;
            }
        }

        private void FileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                if (e.AddedItems[0] is TemplateFileViewModel file)
                {
                    ViewModel?.OpenTemplateFileCommand.Execute(file.FilePath);
                }
                else if (e.AddedItems[0] is string path)
                {
                    ViewModel?.OpenTemplateFileCommand.Execute(path);
                }
                ((ComboBox)sender).SelectedItem = null;
            }
        }

        private void OnManagePromptRequested(object? sender, EventArgs e)
        {
            System.Windows.MessageBox.Show(Properties.Resources.Prompt_management_interface_no_c1825d, "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        private void RangeButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                btn.ContextMenu.IsOpen = true;
            }
        }

        private void MailGrid_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel is null)
            {
                return;
            }

            // Cancel any pending preview to avoid interference with double-click action
            ViewModel.CancelPreview();

            if (sender is DataGrid grid && grid.SelectedItem is MailItemViewModel mail)
            {
                var application = AddInContext.OutlookApplication;
                if (application is null)
                {
                    return;
                }

                MailItem? item = null;
                try
                {
                    item = ResolveMail(application, mail);
                    
                    if (item is null)
                    {
                        // Strategy 1: Try InternetMessageId first (Exact Match)
                        // This fixes the issue where fuzzy subject search finds the wrong email (e.g. Reply instead of Sent)
                        if (!string.IsNullOrEmpty(mail.InternetMessageId))
                        {
                            item = FindMailByInternetMessageId(application, mail.InternetMessageId);
                        }

                        // Strategy 2: Fallback to Original Subject Search
                        if (item is null)
                        {
                            item = FindMailBySubject(application, mail.Subject);
                        }

                        if (item is not null)
                        {
                            item.Display();
                            
                            var newEntryId = item.EntryID;
                            var newStoreId = ((MAPIFolder)item.Parent).StoreID;
                            var oldEntryId = mail.EntryId;
                            var eventId = ViewModel?.EventId;
                            var repo = ViewModel?.Services?.EventRepository;

                            if (!string.IsNullOrEmpty(eventId) && repo != null)
                            {
                                _ = Task.Run(async () => 
                                {
                                    try 
                                    {
                                        var record = await repo.GetByIdAsync(eventId!);
                                        if (record != null)
                                        {
                                            var email = record.Emails.FirstOrDefault(e => e.EntryId == oldEntryId);
                                            if (email != null)
                                            {
                                                email.EntryId = newEntryId;
                                                email.StoreId = newStoreId;
                                                await repo.UpdateAsync(record);
                                            }
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        Debug.WriteLine($"[EventDetailView] Failed to update mail EntryID: {ex.Message}");
                                    }
                                });
                            }
                            
                            mail.IsNewOrUpdated = false;
                            return;
                        }
                    }

                    if (item is null)
                    {
                        MessageBox.Show(Properties.Resources.Outlook_cannot_find_this_email_9e5a34, Properties.Resources.Email_Unavailable, MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    item.Display();
                    mail.IsNewOrUpdated = false;
                }
                catch (COMException ex)
                {
                    var errorCode = (uint)ex.ErrorCode;
                    Debug.WriteLine($"[EventDetailView] Failed to open mail {mail.EntryId}: 0x{errorCode:X8} {ex.Message}");
                    var message = errorCode == 0x8004010Fu
                        ? Properties.Resources.Outlook_cannot_find_this_email_9e5a34
                        : Properties.Resources.Error_opening_email_please_try_again_later;
                    MessageBox.Show(message, Properties.Resources.Cannot_Open_Email, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                finally
                {
                    if (item is not null)
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
        }

        private void MailGrid_OnDragOver(object sender, DragEventArgs e)
        {
            var command = ViewModel?.HandleMailDropCommand;
            e.Effects = command?.CanExecute(null) == true ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void MailGrid_OnDrop(object sender, DragEventArgs e)
        {
            var command = ViewModel?.HandleMailDropCommand;
            if (command?.CanExecute(null) == true)
            {
                command.Execute(null);
            }

            e.Handled = true;
        }

        private void OnLlmSettingsRequested(object? sender, EventArgs e)
        {
            if (ViewModel is null)
            {
                return;
            }

            var services = ViewModel.Services;
            var templateId = ViewModel.SelectedTemplate?.TemplateId;
            var vm = new LlmSettingsViewModel(services.LlmConfigurations, services.OllamaModels, templateId);
            var window = new LlmSettingsWindow(vm)
            {
                Owner = Window.GetWindow(this)
            };

            window.ShowDialog();
        }

        private static MailItem? ResolveMail(Microsoft.Office.Interop.Outlook.Application application, MailItemViewModel mail)
        {
            if (application is null || mail is null)
            {
                return null;
            }

            MailItem? item = null;
            var session = application.Session;

            if (!string.IsNullOrEmpty(mail.StoreId))
            {
                try
                {
                    item = session.GetItemFromID(mail.EntryId, mail.StoreId) as MailItem;
                }
                catch (COMException)
                {
                    item = null;
                }
                catch
                {
                    item = null;
                }
            }

            if (item is null)
            {
                try
                {
                    item = session.GetItemFromID(mail.EntryId) as MailItem;
                }
                catch (COMException)
                {
                    item = null;
                }
                catch
                {
                    item = null;
                }
            }

            return item;
        }

        private static MailItem? FindMailBySubject(Microsoft.Office.Interop.Outlook.Application application, string subject)
        {
            if (string.IsNullOrEmpty(subject)) return null;

            try
            {
                // Use LIKE for fuzzy matching to handle prefixes like RE:, FW:, (URGENT) etc.
                string filter = $"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject.Replace("'", "''")}%'";
                
                // Search Inbox
                var inbox = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                var item = FindMailRecursive(inbox, filter, application.Session);
                if (item != null) return item;

                // Search Sent Items
                var sent = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
                return FindMailRecursive(sent, filter, application.Session);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"[EventDetailView] FindMailBySubject failed: {ex.Message}");
                return null;
            }
        }

        private static MailItem? FindMailRecursive(MAPIFolder folder, string filter, NameSpace session)
        {
            try
            {
                var table = folder.GetTable(filter, OlTableContents.olUserItems);
                if (table.GetRowCount() > 0)
                {
                    var row = table.GetNextRow();
                    var entryId = row["EntryID"] as string;
                    if (!string.IsNullOrEmpty(entryId))
                    {
                        return session.GetItemFromID(entryId) as MailItem;
                    }
                }

                foreach (MAPIFolder subFolder in folder.Folders)
                {
                    var item = FindMailRecursive(subFolder, filter, session);
                    if (item != null) return item;
                }
            }
            catch { }

            return null;
        }

        private static MailItem? FindMailByInternetMessageId(Microsoft.Office.Interop.Outlook.Application application, string internetMessageId)
        {
            if (string.IsNullOrEmpty(internetMessageId)) return null;

            // DASL query for PR_INTERNET_MESSAGE_ID
            string filter = $"@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x1035001F\" = '{internetMessageId.Replace("'", "''")}'";

            var session = application.Session;

            // Search Inbox
            var inbox = session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            var item = FindMailInFolder(inbox, filter);
            if (item != null) return item;

            // Search Sent Items
            var sent = session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            item = FindMailInFolder(sent, filter);
            if (item != null) return item;

            // Search Deleted Items
            var deleted = session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            return FindMailInFolder(deleted, filter);
        }

        private static MailItem? FindMailInFolder(MAPIFolder folder, string filter)
        {
            if (folder == null) return null;
            
            Table? table = null;
            try
            {
                table = folder.GetTable(filter, OlTableContents.olUserItems);
                if (table.GetRowCount() > 0)
                {
                    var row = table.GetNextRow();
                    var entryId = row["EntryID"] as string;
                    if (!string.IsNullOrEmpty(entryId))
                    {
                        return folder.Session.GetItemFromID(entryId) as MailItem;
                    }
                }
            }
            catch { }
            finally
            {
                if (table != null) Marshal.ReleaseComObject(table);
                if (folder != null) Marshal.ReleaseComObject(folder);
            }
            return null;
        }

        private Point _startPoint;
        private bool _isDragging;
        private bool _pendingSelectionReset;
        private object? _pendingItem;

        private static ListViewItem? GetListViewItem(DependencyObject? obj)
        {
            while (obj != null && obj is not ListViewItem)
            {
                obj = System.Windows.Media.VisualTreeHelper.GetParent(obj);
            }
            return obj as ListViewItem;
        }

        private void HandlePreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);

            var listView = sender as ListView;
            if (listView == null) return;

            var item = GetListViewItem(e.OriginalSource as DependencyObject);
            if (item != null && item.IsSelected && (Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                // If clicking an already selected item without Ctrl, we might be starting a drag.
                // Swallow the event to prevent immediate selection reset.
                e.Handled = true;
                _pendingSelectionReset = true;
                _pendingItem = item.DataContext;
            }
            else
            {
                _pendingSelectionReset = false;
                _pendingItem = null;
            }
        }

        private void HandlePreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_pendingSelectionReset && _pendingItem != null)
            {
                // User clicked a selected item but didn't drag.
                // Reset selection to just this item.
                var listView = sender as ListView;
                if (listView != null)
                {
                    listView.SelectedItems.Clear();
                    
                    // Find the item container to select it
                    var itemContainer = listView.ItemContainerGenerator.ContainerFromItem(_pendingItem) as ListViewItem;
                    if (itemContainer != null)
                    {
                        itemContainer.IsSelected = true;
                        itemContainer.Focus();
                    }
                    else 
                    {
                        // Fallback if container not found (e.g. virtualized)
                        // Try to add to SelectedItems directly if possible, or use SelectedItem
                        listView.SelectedItem = _pendingItem;
                    }
                }
                _pendingSelectionReset = false;
                _pendingItem = null;
            }
        }

        private void AttachmentList_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount >= 2)
            {
                AttachmentList_OnMouseDoubleClick(sender, e);
                e.Handled = true;
                return;
            }
            HandlePreviewMouseLeftButtonDown(sender, e);
        }

        private void AttachmentList_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            HandlePreviewMouseLeftButtonUp(sender, e);
        }

        private void AttachmentList_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point position = e.GetPosition(null);
                if (Math.Abs(position.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // If we start dragging, cancel the pending reset
                    _pendingSelectionReset = false;
                    _pendingItem = null;
                    StartDragAttachments(sender as ListView);
                }
            }
        }

        private void StartDragAttachments(ListView? list)
        {
            if (list is null || list.SelectedItems.Count == 0) return;

            _isDragging = true;
            try
            {
                var itemsToProcess = new System.Collections.Generic.List<AttachmentItemViewModel>();
                foreach (var item in list.SelectedItems)
                {
                    if (item is AttachmentItemViewModel vm)
                    {
                        itemsToProcess.Add(vm);
                    }
                }

                var paths = new System.Collections.Generic.List<string>();
                foreach (var attachmentVm in itemsToProcess)
                {
                    // Synchronous extraction to ensure DoDragDrop works
                    var path = ExtractAttachment(attachmentVm);
                    if (!string.IsNullOrEmpty(path))
                    {
                        paths.Add(path!);
                    }
                }

                if (paths.Count > 0)
                {
                    var data = new DataObject(DataFormats.FileDrop, paths.ToArray());
                    DragDrop.DoDragDrop(list, data, DragDropEffects.Copy);
                }
            }
            finally
            {
                _isDragging = false;
            }
        }

        private void AttachmentList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            AttachmentItemViewModel? attachmentVm = null;
            var listViewItem = GetListViewItem(e.OriginalSource as DependencyObject);
            if (listViewItem != null)
            {
                attachmentVm = listViewItem.DataContext as AttachmentItemViewModel;
            }
            else if (sender is ListView list)
            {
                attachmentVm = list.SelectedItem as AttachmentItemViewModel;
            }

            if (attachmentVm != null)
            {
                // Execute on UI thread to avoid COM threading issues (RPC_E_WRONG_THREAD)
                var path = ExtractAttachment(attachmentVm);
                if (!string.IsNullOrEmpty(path))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(string.Format(Properties.Resources.Cannot_open_file_ex_Message, ex.Message));
                    }
                }
            }
        }

        private string? ExtractAttachment(AttachmentItemViewModel vm)
        {
            try
            {
                var app = AddInContext.OutlookApplication;
                if (app == null) return null;

                var ns = app.GetNamespace("MAPI");
                MailItem? item = null;
                try
                {
                    item = ns.GetItemFromID(vm.SourceMailEntryId) as MailItem;
                }
                catch
                {
                    // Ignore invalid entry ID
                }

                if (item == null) return null;

                try
                {
                    // Try exact match first (name + size)
                    foreach (Attachment attachment in item.Attachments)
                    {
                        if (string.Equals(attachment.FileName, vm.FileName, StringComparison.OrdinalIgnoreCase) && 
                            attachment.Size == vm.FileSizeBytes)
                        {
                            var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), vm.FileName);
                            attachment.SaveAsFile(tempPath);
                            return tempPath;
                        }
                    }

                    // Fallback: try matching just filename
                    foreach (Attachment attachment in item.Attachments)
                    {
                        if (string.Equals(attachment.FileName, vm.FileName, StringComparison.OrdinalIgnoreCase))
                        {
                            var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), vm.FileName);
                            attachment.SaveAsFile(tempPath);
                            return tempPath;
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"Failed to extract attachment: {ex.Message}");
            }
            return null;
        }

        private void TemplateFileList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TemplateFileViewModel? file = null;
            var listViewItem = GetListViewItem(e.OriginalSource as DependencyObject);
            if (listViewItem != null)
            {
                file = listViewItem.DataContext as TemplateFileViewModel;
            }
            else if (sender is ListView list)
            {
                file = list.SelectedItem as TemplateFileViewModel;
            }

            if (file != null)
            {
                ViewModel?.OpenTemplateFileCommand.Execute(file.FilePath);
            }
        }

        private void TemplateFileList_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount >= 2)
            {
                TemplateFileList_OnMouseDoubleClick(sender, e);
                e.Handled = true;
                return;
            }
            HandlePreviewMouseLeftButtonDown(sender, e);
        }

        private void TemplateFileList_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            HandlePreviewMouseLeftButtonUp(sender, e);
        }

        private void TemplateFileList_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point position = e.GetPosition(null);
                if (Math.Abs(position.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // If we start dragging, cancel the pending reset
                    _pendingSelectionReset = false;
                    _pendingItem = null;
                    StartDragTemplateFiles(sender as ListView);
                }
            }
        }

        private void StartDragTemplateFiles(ListView? list)
        {
            if (list is null || list.SelectedItems.Count == 0) return;

            _isDragging = true;
            try
            {
                var paths = list.SelectedItems.Cast<TemplateFileViewModel>()
                                .Select(f => f.FilePath)
                                .Where(p => !string.IsNullOrEmpty(p) && System.IO.File.Exists(p))
                                .ToArray();

                if (paths.Length > 0)
                {
                    var data = new DataObject(DataFormats.FileDrop, paths);
                    DragDrop.DoDragDrop(list, data, DragDropEffects.Copy);
                }
            }
            finally
            {
                _isDragging = false;
            }
        }

        private void TemplateFileList_OnDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                ViewModel?.AddTemplateFiles(files);
            }
        }
    }
}
