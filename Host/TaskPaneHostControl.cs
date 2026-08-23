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
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
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

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Chuyển tiếp các phím tắt chỉnh sửa văn bản (Ctrl+X Cut, Ctrl+C Copy, Ctrl+V Paste, Ctrl+A SelectAll, Ctrl+Z Undo, Ctrl+Y Redo)
            // tới ô TextBox/PasswordBox đang focus trong TaskPane WPF thay vì để Excel chặn bắt
            if (keyData == (Keys.Control | Keys.X) ||
                keyData == (Keys.Control | Keys.C) ||
                keyData == (Keys.Control | Keys.V) ||
                keyData == (Keys.Control | Keys.A) ||
                keyData == (Keys.Control | Keys.Z) ||
                keyData == (Keys.Control | Keys.Y) ||
                keyData == (Keys.Shift | Keys.Delete) ||
                keyData == (Keys.Control | Keys.Insert) ||
                keyData == (Keys.Shift | Keys.Insert))
            {
                if (System.Windows.Input.Keyboard.FocusedElement is System.Windows.Controls.TextBox tb)
                {
                    if (keyData == (Keys.Control | Keys.X) || keyData == (Keys.Shift | Keys.Delete))
                    {
                        tb.Cut();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.C) || keyData == (Keys.Control | Keys.Insert))
                    {
                        tb.Copy();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.V) || keyData == (Keys.Shift | Keys.Insert))
                    {
                        tb.Paste();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.A))
                    {
                        tb.SelectAll();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.Z))
                    {
                        if (tb.CanUndo) tb.Undo();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.Y))
                    {
                        if (tb.CanRedo) tb.Redo();
                        return true;
                    }
                }
                else if (System.Windows.Input.Keyboard.FocusedElement is System.Windows.Controls.PasswordBox pb)
                {
                    if (keyData == (Keys.Control | Keys.V) || keyData == (Keys.Shift | Keys.Insert))
                    {
                        pb.Paste();
                        return true;
                    }
                    if (keyData == (Keys.Control | Keys.A))
                    {
                        pb.SelectAll();
                        return true;
                    }
                }
            }

            return base.ProcessCmdKey(ref msg, keyData);
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
