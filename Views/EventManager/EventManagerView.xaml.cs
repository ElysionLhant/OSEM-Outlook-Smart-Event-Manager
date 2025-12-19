#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using OSEMAddIn.ViewModels;
using OSEMAddIn.Views;

namespace OSEMAddIn.Views.EventManager
{
    public partial class EventManagerView : UserControl
    {
        public EventManagerView()
        {
            InitializeComponent();
        }

        private EventManagerViewModel? ViewModel => DataContext as EventManagerViewModel;

        private void OngoingGrid_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ViewModel is null || sender is not DataGrid grid)
            {
                return;
            }

            var items = grid.SelectedItems.Cast<object>().ToList();
            ViewModel.UpdateSelection(items);
        }

        private void OngoingGrid_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel is null)
            {
                return;
            }

            // Prevent double click if the source is the priority button
            if (e.OriginalSource is DependencyObject dep)
            {
                // Traverse up the visual tree to see if we clicked a Button
                var parent = dep;
                while (parent != null && parent is not DataGridRow && parent is not DataGrid)
                {
                    if (parent is Button)
                    {
                        return;
                    }
                    parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
                }
            }

            if (sender is DataGrid grid && grid.SelectedItem is EventListItemViewModel item)
            {
                ViewModel.OpenEventCommand.Execute(item);
            }
        }

        private void ArchivedGrid_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel is null)
            {
                return;
            }

            if (sender is DataGrid grid && grid.SelectedItem is EventListItemViewModel item)
            {
                ViewModel.OpenEventCommand.Execute(item);
            }
        }

        private void OngoingGrid_OnDragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void OngoingGrid_OnDrop(object sender, DragEventArgs e)
        {
            if (ViewModel is null)
            {
                return;
            }

            ViewModel.HandleMailDropCommand.Execute(null);
            e.Handled = true;
        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = this.ViewModel;
            if (viewModel is null)
            {
                return;
            }

            // Create ViewModel for the dialog
            var templates = viewModel.DashboardTemplates;
            var initialTemplate = viewModel.SelectedFilterTemplate;
            var startDate = viewModel.FilterStartDate ?? System.DateTime.Now.AddMonths(-1);
            var endDate = viewModel.FilterEndDate ?? System.DateTime.Now;

            var exportVm = new ExportOptionsViewModel(templates, initialTemplate, startDate, endDate);
            
            // Create Window
            var window = new ExportOptionsWindow
            {
                DataContext = exportVm,
                Owner = Window.GetWindow(this)
            };

            // Handle Browse
            exportVm.BrowseRequested += () =>
            {
                var dialog = new System.Windows.Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    exportVm.TargetPath = dialog.SelectedPath;
                }
            };

            // Handle Close/Export
            exportVm.RequestClose += async (shouldExport) =>
            {
                window.Close();
                if (shouldExport)
                {
                    await viewModel.ExportEventsAsync(exportVm);
                }
            };

            window.Show();
        }

        private void OngoingGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            if (ViewModel is null) return;

            if (e.Column.SortMemberPath == "PriorityLevel")
            {
                e.Handled = true;

                using (ViewModel.OngoingEvents.DeferRefresh())
                {
                    ViewModel.OngoingEvents.SortDescriptions.Clear();

                    if (e.Column.SortDirection == null)
                    {
                        // 1. First click: Descending (Red -> Orange -> Yellow -> None)
                        e.Column.SortDirection = System.ComponentModel.ListSortDirection.Descending;
                        ViewModel.OngoingEvents.SortDescriptions.Add(new System.ComponentModel.SortDescription("PriorityLevel", System.ComponentModel.ListSortDirection.Descending));
                    }
                    else if (e.Column.SortDirection == System.ComponentModel.ListSortDirection.Descending)
                    {
                        // 2. Second click: Ascending (None -> Yellow -> Orange -> Red)
                        e.Column.SortDirection = System.ComponentModel.ListSortDirection.Ascending;
                        ViewModel.OngoingEvents.SortDescriptions.Add(new System.ComponentModel.SortDescription("PriorityLevel", System.ComponentModel.ListSortDirection.Ascending));
                    }
                    else
                    {
                        // 3. Third click: Clear sorting (Original order)
                        e.Column.SortDirection = null;
                    }
                }
            }
        }

        private void DataGrid_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (ViewModel == null || sender is not DataGrid grid) return;

            // Restore SortDirection from CollectionView or Saved State
            System.ComponentModel.ICollectionView? view = null;
            System.ComponentModel.SortDescription? savedSort = null;

            if (grid == OngoingGrid)
            {
                view = ViewModel.OngoingEvents;
                savedSort = ViewModel.SavedOngoingSort;
            }
            else if (grid == ArchivedGrid)
            {
                view = ViewModel.ArchivedEvents;
                savedSort = ViewModel.SavedArchivedSort;
            }

            if (view != null)
            {
                // If CollectionView lost its sort but we have a saved one, restore it
                if (view.SortDescriptions.Count == 0 && savedSort.HasValue)
                {
                    view.SortDescriptions.Add(savedSort.Value);
                }

                // Sync DataGrid column headers with CollectionView sort
                if (view.SortDescriptions.Count > 0)
                {
                    var sortDesc = view.SortDescriptions[0];
                    var column = grid.Columns.FirstOrDefault(c => 
                        c.SortMemberPath == sortDesc.PropertyName || 
                        (c is DataGridBoundColumn bound && (bound.Binding as System.Windows.Data.Binding)?.Path?.Path == sortDesc.PropertyName));
                    
                    if (column != null)
                    {
                        column.SortDirection = sortDesc.Direction;
                    }
                }
            }

            double offset = 0;
            if (grid == OngoingGrid)
            {
                offset = ViewModel.OngoingScrollOffset;
            }
            else if (grid == ArchivedGrid)
            {
                offset = ViewModel.ArchivedScrollOffset;
            }

            // Try to restore immediately to avoid visual jump
            if (offset > 0)
            {
                var scrollViewer = GetScrollViewer(grid);
                if (scrollViewer != null)
                {
                    scrollViewer.ScrollToVerticalOffset(offset);
                }
                else
                {
                    // Fallback if ScrollViewer not ready yet
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        GetScrollViewer(grid)?.ScrollToVerticalOffset(offset);
                    }), System.Windows.Threading.DispatcherPriority.Render);
                }
            }

            // Delay subscription to avoid capturing initial layout changes
            Dispatcher.BeginInvoke(new Action(() =>
            {
                grid.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(DataGrid_OnScrollChanged));
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        private void DataGrid_OnUnloaded(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                grid.RemoveHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(DataGrid_OnScrollChanged));
            }
        }

        private void DataGrid_OnScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (ViewModel == null || sender is not DataGrid grid) return;
            
            // Only handle scroll events from the DataGrid's main ScrollViewer
            if (e.OriginalSource is ScrollViewer scrollViewer && scrollViewer.TemplatedParent == grid)
            {
                if (grid == OngoingGrid)
                {
                    ViewModel.OngoingScrollOffset = scrollViewer.VerticalOffset;
                }
                else if (grid == ArchivedGrid)
                {
                    ViewModel.ArchivedScrollOffset = scrollViewer.VerticalOffset;
                }
            }
        }

        private ScrollViewer? GetScrollViewer(DependencyObject depObj)
        {
            if (depObj is ScrollViewer scrollViewer) return scrollViewer;

            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(depObj, i);
                var result = GetScrollViewer(child);
                if (result != null) return result;
            }
            return null;
        }
    }
}
