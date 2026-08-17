using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using ExcelSupport.Models;
using ExcelSupport.Services;
using MessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class GlossaryDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private readonly ObservableCollection<GlossaryItem> _items = new ObservableCollection<GlossaryItem>();
        private readonly ICollectionView _view;

        public bool IsDarkTheme
        {
            get => _isDarkTheme;
            set
            {
                if (_isDarkTheme != value)
                {
                    _isDarkTheme = value;
                    OnPropertyChanged(nameof(IsDarkTheme));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public GlossaryDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            // Load existing glossary
            var existing = AiConfigManager.Current.Glossary ?? new List<GlossaryItem>();
            foreach (var item in existing)
            {
                _items.Add(item.Clone());
            }

            _view = CollectionViewSource.GetDefaultView(_items);
            _view.Filter = FilterGlossary;
            GridGlossary.ItemsSource = _view;

            UpdateCountBadge();
        }

        private bool FilterGlossary(object obj)
        {
            if (obj is GlossaryItem item)
            {
                string search = TxtSearch.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(search)) return true;

                return (!string.IsNullOrEmpty(item.Japanese) && item.Japanese.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.Vietnamese) && item.Vietnamese.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (!string.IsNullOrEmpty(item.Note) && item.Note.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return true;
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _view?.Refresh();
            UpdateCountBadge();
        }

        private void UpdateCountBadge()
        {
            int total = _items.Count;
            TxtTotalCount.Text = $"{total} thuật ngữ";
        }

        private void OnAddRowClick(object sender, RoutedEventArgs e)
        {
            var newItem = new GlossaryItem
            {
                Japanese = string.Empty,
                Vietnamese = string.Empty,
                Note = string.Empty
            };
            _items.Insert(0, newItem);
            GridGlossary.SelectedItem = newItem;
            GridGlossary.ScrollIntoView(newItem);
            UpdateCountBadge();
            TxtStatus.Text = "➕ Đã thêm dòng mới. Vui lòng nhập từ tiếng Nhật và tiếng Việt vào bảng.";
        }

        private void OnDeleteRowClick(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is GlossaryItem item)
            {
                _items.Remove(item);
                UpdateCountBadge();
                TxtStatus.Text = $"🗑️ Đã xóa thuật ngữ: [{item.Japanese} ⇋ {item.Vietnamese}]";
            }
        }

        private void OnClearAllClick(object sender, RoutedEventArgs e)
        {
            if (_items.Count == 0) return;

            var confirm = MessageBox.Show("Bạn có chắc chắn muốn xóa toàn bộ danh sách thuật ngữ không?",
                                          "Xác nhận xóa hết", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm == MessageBoxResult.Yes)
            {
                _items.Clear();
                UpdateCountBadge();
                TxtStatus.Text = "🗑️ Đã xóa toàn bộ thuật ngữ.";
            }
        }

        private void OnImportClick(object sender, RoutedEventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Chọn file Glossary để nhập (CSV hoặc JSON)";
                dlg.Filter = "Tất cả file hỗ trợ (*.csv;*.json)|*.csv;*.json|File CSV (*.csv)|*.csv|File JSON (*.json)|*.json|Tất cả file (*.*)|*.*";

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        string ext = System.IO.Path.GetExtension(dlg.FileName).ToLowerInvariant();
                        List<GlossaryItem> importedList;

                        if (ext == ".json")
                        {
                            importedList = GlossaryService.ImportFromJson(dlg.FileName);
                        }
                        else
                        {
                            importedList = GlossaryService.ImportFromCsv(dlg.FileName);
                        }

                        if (importedList.Count == 0)
                        {
                            MessageBox.Show("Không tìm thấy dữ liệu thuật ngữ hợp lệ trong file.", "Import Glossary", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        // Hỏi người dùng muốn ghi đè hay gộp thêm
                        var choice = MessageBox.Show($"Tìm thấy {importedList.Count} thuật ngữ trong file.\n\n" +
                                                     $"Bấm [Yes] để GỘP THÊM vào danh sách hiện tại.\n" +
                                                     $"Bấm [No] để THAY THẾ TOÀN BỘ danh sách hiện tại.",
                                                     "Tùy chọn Import", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                        if (choice == MessageBoxResult.Cancel) return;

                        if (choice == MessageBoxResult.No)
                        {
                            _items.Clear();
                        }

                        int addedCount = 0;
                        foreach (var item in importedList)
                        {
                            // Kiểm tra trùng lặp
                            bool exists = _items.Any(i => string.Equals(i.Japanese?.Trim(), item.Japanese?.Trim(), StringComparison.OrdinalIgnoreCase));
                            if (!exists)
                            {
                                _items.Add(item);
                                addedCount++;
                            }
                        }

                        UpdateCountBadge();
                        TxtStatus.Text = $"✅ Đã nhập thành công {addedCount} thuật ngữ mới từ file!";
                        MessageBox.Show($"✅ Đã nhập thành công {addedCount} thuật ngữ mới!", "Import Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi nhập file:\n{ex.Message}", "Lỗi Import", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            if (_items.Count == 0)
            {
                MessageBox.Show("Danh sách Glossary đang trống, không có dữ liệu để xuất.", "Export Glossary", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Lưu file Glossary xuất ra";
                dlg.Filter = "File CSV (*.csv)|*.csv|File JSON (*.json)|*.json";
                dlg.FileName = $"Glossary_ExcelSupport_{DateTime.Now:yyyyMMdd}";

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        string ext = System.IO.Path.GetExtension(dlg.FileName).ToLowerInvariant();
                        if (ext == ".json")
                        {
                            GlossaryService.ExportToJson(dlg.FileName, _items);
                        }
                        else
                        {
                            GlossaryService.ExportToCsv(dlg.FileName, _items);
                        }

                        TxtStatus.Text = $"✅ Đã xuất thành công {_items.Count} thuật ngữ vào file: {System.IO.Path.GetFileName(dlg.FileName)}";
                        MessageBox.Show($"✅ Đã xuất {_items.Count} thuật ngữ ra file:\n{dlg.FileName}", "Export Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi xuất file:\n{ex.Message}", "Lỗi Export", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void OnSaveAndCloseClick(object sender, RoutedEventArgs e)
        {
            // Lọc các dòng hợp lệ
            var validItems = _items.Where(i => !string.IsNullOrWhiteSpace(i.Japanese) || !string.IsNullOrWhiteSpace(i.Vietnamese)).ToList();

            var config = AiConfigManager.Current;
            config.Glossary = validItems;
            AiConfigManager.Save(config);

            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
