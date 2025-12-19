using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using OSEMAddIn.Services;
using OSEMAddIn.ViewModels;
using OSEMAddIn.Views.Shell;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OSEMAddIn
{
    public partial class ThisAddIn
    {
        private ServiceContainer? _services;
        private ShellView? _shellView;
    private ElementHost? _elementHost;
    private UserControl? _taskPaneContainer;
        private Microsoft.Office.Tools.CustomTaskPane? _taskPane;
        private Outlook.SyncObjects? _syncObjects;
        private int _syncCount = 0;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (SynchronizationContext.Current is null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            AddInContext.Initialize(Application);

            _services = new ServiceContainer(Application);
            _services.EventMonitor.Start();
            
            SetupSyncMonitoring();

            var shellViewModel = new ShellViewModel(_services);

            _shellView = new ShellView
            {
                DataContext = shellViewModel
            };

            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _shellView
            };

            _taskPaneContainer = new UserControl
            {
                Dock = DockStyle.Fill
            };
            _taskPaneContainer.Controls.Add(_elementHost);

            _taskPane = CustomTaskPanes.Add(_taskPaneContainer, Properties.Resources.Event_Manager);
            _taskPane.Width = 420;
            _taskPane.Visible = true;
        }

        private void SetupSyncMonitoring()
        {
            try
            {
                _syncObjects = Application.Session.SyncObjects;
                for (int i = 1; i <= _syncObjects.Count; i++)
                {
                    var syncObject = _syncObjects[i];
                    syncObject.SyncStart += SyncObject_SyncStart;
                    syncObject.SyncEnd += SyncObject_SyncEnd;
                    syncObject.OnError += SyncObject_OnError;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting up sync monitoring: {ex}");
            }
        }

        private void SyncObject_SyncStart()
        {
            Interlocked.Increment(ref _syncCount);
            UpdateBusyState();
        }

        private void SyncObject_SyncEnd()
        {
            Interlocked.Decrement(ref _syncCount);
            UpdateBusyState();
        }

        private void SyncObject_OnError(int Code, string Description)
        {
            // Just log or ignore, SyncEnd should fire eventually
            // Interlocked.Decrement(ref _syncCount); 
            // UpdateBusyState();
        }

        private void UpdateBusyState()
        {
            if (_shellView != null && _shellView.Dispatcher != null)
            {
                _shellView.Dispatcher.InvokeAsync(() =>
                {
                    if (_syncCount > 0)
                        _services?.BusyState.SetBusy("Outlook is synchronizing...");
                    else
                        _services?.BusyState.ClearBusy();
                });
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_syncObjects != null)
            {
                for (int i = 1; i <= _syncObjects.Count; i++)
                {
                    var syncObject = _syncObjects[i];
                    syncObject.SyncStart -= SyncObject_SyncStart;
                    syncObject.SyncEnd -= SyncObject_SyncEnd;
                    syncObject.OnError -= SyncObject_OnError;
                }
                Marshal.ReleaseComObject(_syncObjects);
                _syncObjects = null;
            }

            _taskPane?.Dispose();
            _elementHost?.Dispose();
            _taskPaneContainer?.Dispose();
            _shellView = null;

            if (_services is not null)
            {
                _services.EventMonitor.Stop();
                _services.Dispose();
            }
            _services = null;

            AddInContext.Reset();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
