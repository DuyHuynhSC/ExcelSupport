using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelSupport.ViewModels;
using ExcelSupport.Views;
using WpfApp = System.Windows.Application;
using WpfShutdownMode = System.Windows.ShutdownMode;

namespace ExcelSupport.Host
{
    [ComVisible(true)]
    [Guid("E3C2D819-5432-47C2-88E2-6E3C0F55A8A4")]
    [ProgId("ExcelSupport.TaskPaneHostControl")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneHostControl : UserControl
    {
        public ElementHost? Host { get; private set; }
        public WorkbookTreeViewControl? WpfView { get; private set; }

        public TaskPaneHostControl()
        {
            // Giữ constructor tối giản tuyệt đối để COM CoCreateInstance luôn thành công
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            InitializeBridge();
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (Visible)
            {
                Host?.Focus();
                AddInEvents.Instance?.QueueRefresh();
            }
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            Host?.Focus();
        }

        private void InitializeBridge()
        {
            if (Host != null) return;

            try
            {
                // Khởi tạo WPF Application runtime nếu chưa có
                if (WpfApp.Current == null)
                {
                    new WpfApp
                    {
                        ShutdownMode = WpfShutdownMode.OnExplicitShutdown
                    };
                }

                Host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    TabStop = true
                };

                WpfView = new WorkbookTreeViewControl
                {
                    DataContext = AddInEvents.MainViewModel
                };

                Host.Child = WpfView;
                Controls.Add(Host);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khởi tạo giao diện WPF trong TaskPane: {ex}",
                                "Lỗi khởi tạo UI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void BindViewModel(TaskPaneViewModel? viewModel)
        {
            if (WpfView != null && viewModel != null)
            {
                WpfView.DataContext = viewModel;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Host?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
