using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfMessageBoxResult = System.Windows.MessageBoxResult;
using MediaColor = System.Windows.Media.Color;
using WpfWindow = System.Windows.Window;
using WpfButton = System.Windows.Controls.Button;

namespace ExcelSupport.Views
{
    public partial class SheetSnapshotDialog : WpfWindow
    {
        private SheetSnapshotItem? _selectedSnapshot;
        private readonly bool _isDarkTheme;
        private static SheetSnapshotDialog? _currentInstance;

        public SheetSnapshotDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            _isDarkTheme = isDarkTheme;

            if (_isDarkTheme)
            {
                ApplyDarkTheme();
            }

            Loaded += SheetSnapshotDialog_Loaded;
        }

        private void SheetSnapshotDialog_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshSnapshotList();
        }

        public static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new SheetSnapshotDialog(isDarkTheme);

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new System.Windows.Interop.WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở Quản lý Snapshot:\n{ex.Message}", "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void ApplyDarkTheme()
        {
            RootGrid.Background = new System.Windows.Media.SolidColorBrush(MediaColor.FromRgb(15, 23, 42));
        }

        private void RefreshSnapshotList()
        {
            var list = SheetSnapshotService.GetSnapshots();
            lbSnapshots.ItemsSource = list;
            txtSnapshotCount.Text = $"{list.Count} bản lưu";

            if (list.Count > 0)
            {
                if (_selectedSnapshot == null || !list.Any(s => s.Id == _selectedSnapshot.Id))
                {
                    lbSnapshots.SelectedIndex = 0;
                }
            }
            else
            {
                ClearSelection();
            }
        }

        private void BtnTakeSnapshot_Click(object sender, RoutedEventArgs e)
        {
            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null)
            {
                txtStatus.Text = "Không thể kết nối với tiến trình Excel.";
                return;
            }

            string desc = txtSnapshotDesc.Text.Trim();
            var snap = SheetSnapshotService.TakeSnapshot(app, desc, isAuto: false);

            if (snap != null)
            {
                txtSnapshotDesc.Text = string.Empty;
                txtStatus.Text = $"Đã chụp snapshot thành công cho Sheet '{snap.SheetName}' ({snap.RowCount:N0} × {snap.ColumnCount:N0} ô).";
                RefreshSnapshotList();
                lbSnapshots.SelectedItem = snap;
            }
            else
            {
                WpfMessageBox.Show("Không thể chụp snapshot (Worksheet trống hoặc đang ở chế độ chỉnh sửa ô).", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
            }
        }

        private void LbSnapshots_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;

            _selectedSnapshot = lbSnapshots.SelectedItem as SheetSnapshotItem;
            if (_selectedSnapshot == null)
            {
                ClearSelection();
                return;
            }

            txtSelectedTitle.Text = $"Bản lưu: {_selectedSnapshot.SheetName} ({_selectedSnapshot.DisplayTimestamp})";

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app != null)
            {
                var diffs = SheetSnapshotService.CompareWithCurrent(app, _selectedSnapshot);
                dgDiffPreview.ItemsSource = diffs;

                if (diffs.Count == 0)
                {
                    txtDiffSummary.Text = "✅ Sheet hiện tại trùng khớp 100% với bản lưu này.";
                }
                else
                {
                    txtDiffSummary.Text = $"⚠️ Phát hiện {diffs.Count:N0} ô có sự thay đổi giữa bản lưu và Sheet hiện tại.";
                }
            }
        }

        private void BtnRollbackCurrent_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedSnapshot == null)
            {
                WpfMessageBox.Show("Vui lòng chọn bản lưu cần khôi phục.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            var confirm = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn khôi phục Sheet '{_selectedSnapshot.SheetName}' về thời điểm [{_selectedSnapshot.DisplayTimestamp}] không?\n\nToàn bộ dữ liệu, công thức và định dạng hiện tại sẽ được thay thế bằng bản lưu.",
                "Xác nhận Khôi phục (Rollback)",
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);

            if (confirm != WpfMessageBoxResult.Yes) return;

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return;

            bool ok = SheetSnapshotService.RestoreSnapshot(app, _selectedSnapshot, restoreToNewSheet: false);
            if (ok)
            {
                txtStatus.Text = $"Đã khôi phục thành công Sheet '{_selectedSnapshot.SheetName}'!";
                WpfMessageBox.Show($"Đã khôi phục thành công Sheet '{_selectedSnapshot.SheetName}' về bản lưu [{_selectedSnapshot.DisplayTimestamp}]!", "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                LbSnapshots_SelectionChanged(this, null!);
            }
            else
            {
                WpfMessageBox.Show("Không thể khôi phục Sheet. Vui lòng kiểm tra lại.", "Lỗi khôi phục", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnRestoreNewSheet_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedSnapshot == null)
            {
                WpfMessageBox.Show("Vui lòng chọn bản lưu cần khôi phục.", "Thông báo", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            var app = AddInEvents.Instance?.ExcelAppInstance;
            if (app == null) return;

            bool ok = SheetSnapshotService.RestoreSnapshot(app, _selectedSnapshot, restoreToNewSheet: true);
            if (ok)
            {
                txtStatus.Text = "Đã xuất bản lưu ra Sheet mới thành công!";
                WpfMessageBox.Show("Đã tạo Sheet mới chứa toàn bộ dữ liệu từ bản lưu thành công!", "Thành công", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            else
            {
                WpfMessageBox.Show("Không thể xuất bản lưu ra Sheet mới.", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnDeleteSnapshot_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Tag is string id)
            {
                SheetSnapshotService.DeleteSnapshot(id);
                RefreshSnapshotList();
            }
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            var confirm = WpfMessageBox.Show("Bạn có chắc chắn muốn xóa toàn bộ danh sách Snapshot đã lưu không?", "Xác nhận", WpfMessageBoxButton.YesNo, WpfMessageBoxImage.Question);
            if (confirm == WpfMessageBoxResult.Yes)
            {
                SheetSnapshotService.ClearAllSnapshots();
                RefreshSnapshotList();
                txtStatus.Text = "Đã xóa toàn bộ bản lưu Snapshot.";
            }
        }

        private void ClearSelection()
        {
            _selectedSnapshot = null;
            txtSelectedTitle.Text = "Vui lòng chọn một bản lưu từ danh sách bên trái.";
            txtDiffSummary.Text = "Chưa chọn bản lưu để so sánh.";
            dgDiffPreview.ItemsSource = null;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
