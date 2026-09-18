using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public class FindReplacePairDisplayItem
    {
        public int Index { get; set; }
        public string FindText { get; set; } = string.Empty;
        public string ReplaceText { get; set; } = string.Empty;
        public int MatchCount { get; set; }
    }

    public partial class BatchFindReplaceDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(BatchFindReplaceDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static BatchFindReplaceDialog? _currentInstance;

        // Lưu giữ trạng thái khi tắt mở lại
        private static string _savedDictionaryText = string.Empty;
        private static FindReplaceScope _savedScope = FindReplaceScope.ActiveSheet;
        private static bool _savedMatchEntireCell = false;
        private static bool _savedMatchCase = false;
        private static bool _savedHighlight = true;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new BatchFindReplaceDialog(app)
                {
                    IsDarkTheme = isDarkTheme
                };

                try
                {
                    if (app != null)
                    {
                        new System.Windows.Interop.WindowInteropHelper(_currentInstance).Owner = (IntPtr)app.Hwnd;
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở màn hình Tìm & Thay Thế Hàng Loạt:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly ExcelApp? _excelApp;
        private List<FindReplacePair> _currentPairs = new List<FindReplacePair>();

        public BatchFindReplaceDialog(ExcelApp? app)
        {
            InitializeComponent();
            _excelApp = app;

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? false;
            }
            catch { }

            Loaded += OnDialogLoaded;
            Closing += (s, e) => SaveState();
        }

        private void OnDialogLoaded(object sender, RoutedEventArgs e)
        {
            // Khôi phục cấu hình trước đó
            if (!string.IsNullOrEmpty(_savedDictionaryText))
            {
                TxtDictionaryInput.Text = _savedDictionaryText;
            }
            else
            {
                // Mẫu gợi ý ban đầu
                TxtDictionaryInput.Text = "Hà Nội => HN\r\nHồ Chí Minh => HCM\r\nĐà Nẵng => ĐN\r\nCông ty TNHH => Cty TNHH";
            }

            switch (_savedScope)
            {
                case FindReplaceScope.Selection: RbScopeSelection.IsChecked = true; break;
                case FindReplaceScope.ActiveSheet: RbScopeActiveSheet.IsChecked = true; break;
                case FindReplaceScope.AllSheetsCurrentWorkbook: RbScopeAllSheets.IsChecked = true; break;
                case FindReplaceScope.AllOpenWorkbooks: RbScopeAllWorkbooks.IsChecked = true; break;
            }

            ChkMatchEntireCell.IsChecked = _savedMatchEntireCell;
            ChkMatchCase.IsChecked = _savedMatchCase;
            ChkHighlight.IsChecked = _savedHighlight;

            RefreshPairsGrid();
        }

        private void SaveState()
        {
            try
            {
                _savedDictionaryText = TxtDictionaryInput.Text;
                _savedScope = GetCurrentScope();
                _savedMatchEntireCell = (ChkMatchEntireCell.IsChecked == true);
                _savedMatchCase = (ChkMatchCase.IsChecked == true);
                _savedHighlight = (ChkHighlight.IsChecked == true);
            }
            catch { }
        }

        private FindReplaceScope GetCurrentScope()
        {
            if (RbScopeSelection.IsChecked == true) return FindReplaceScope.Selection;
            if (RbScopeAllSheets.IsChecked == true) return FindReplaceScope.AllSheetsCurrentWorkbook;
            if (RbScopeAllWorkbooks.IsChecked == true) return FindReplaceScope.AllOpenWorkbooks;
            return FindReplaceScope.ActiveSheet;
        }

        private void OnDictionaryTextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPairsGrid();
        }

        private void RefreshPairsGrid()
        {
            _currentPairs = BatchFindReplaceService.ParseDictionaryText(TxtDictionaryInput.Text);
            TxtPairCountBadge.Text = $"{_currentPairs.Count:N0} cặp từ";

            var displayList = _currentPairs.Select((p, idx) => new FindReplacePairDisplayItem
            {
                Index = idx + 1,
                FindText = p.FindText,
                ReplaceText = p.ReplaceText,
                MatchCount = p.MatchCount
            }).ToList();

            GridPairs.ItemsSource = displayList;
        }

        private void OnLoadFromSelectionClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            try
            {
                var selection = _excelApp.Selection as Range;
                if (selection == null || selection.Rows.Count == 0)
                {
                    WpfMessageBox.Show("Vui lòng bôi đen 1 vùng bảng 2 cột trên Excel (Cột 1: Từ cũ, Cột 2: Từ mới) để nạp vào từ điển.",
                                       "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var loadedPairs = BatchFindReplaceService.LoadDictionaryFromExcelRange(selection);
                if (loadedPairs.Count == 0)
                {
                    WpfMessageBox.Show("Không tìm thấy dữ liệu hợp lệ trong vùng đang chọn.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var lines = loadedPairs.Select(p => $"{p.FindText} => {p.ReplaceText}");
                TxtDictionaryInput.Text = string.Join("\r\n", lines);

                WpfMessageBox.Show($"Đã nạp thành công {loadedPairs.Count:N0} cặp từ tra cứu từ vùng bôi đen trên Excel!",
                                   "Nạp Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi nạp từ vùng chọn:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnPasteClipboardClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (System.Windows.Clipboard.ContainsText())
                {
                    string clipText = System.Windows.Clipboard.GetText();
                    TxtDictionaryInput.Text = clipText;
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể đọc từ Clipboard:\n{ex.Message}", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnClearDictionaryClick(object sender, RoutedEventArgs e)
        {
            TxtDictionaryInput.Text = string.Empty;
        }

        private void OnExecuteReplaceClick(object sender, RoutedEventArgs e)
        {
            if (_excelApp == null) return;

            if (_currentPairs.Count == 0)
            {
                WpfMessageBox.Show("Vui lòng nhập hoặc nạp ít nhất một cặp từ tra cứu để thay thế.", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var options = new BatchFindReplaceOptions
            {
                Scope = GetCurrentScope(),
                MatchEntireCell = (ChkMatchEntireCell.IsChecked == true),
                MatchCase = (ChkMatchCase.IsChecked == true),
                HighlightReplacedCells = (ChkHighlight.IsChecked == true),
                Pairs = _currentPairs
            };

            string scopeDesc = options.Scope switch
            {
                FindReplaceScope.Selection => "vùng đang chọn",
                FindReplaceScope.ActiveSheet => "Sheet hiện tại",
                FindReplaceScope.AllSheetsCurrentWorkbook => "tất cả các Sheet trong file hiện tại",
                FindReplaceScope.AllOpenWorkbooks => "tất cả các file Excel đang mở",
                _ => "Sheet hiện tại"
            };

            var confirm = WpfMessageBox.Show(
                $"Bạn có chắc chắn muốn thực hiện Tìm & Thay Thế {_currentPairs.Count:N0} cặp từ trên {scopeDesc} không?",
                "Xác Nhận Thay Thế Hàng Loạt",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            try
            {
                TxtFooterStatus.Text = "⏳ Đang thực thi tìm & thay thế hàng loạt...";

                var result = BatchFindReplaceService.ExecuteBatchReplace(_excelApp, options);

                // Cập nhật số lần khớp lên Grid
                var displayList = result.PairResults.Select((p, idx) => new FindReplacePairDisplayItem
                {
                    Index = idx + 1,
                    FindText = p.FindText,
                    ReplaceText = p.ReplaceText,
                    MatchCount = p.MatchCount
                }).ToList();

                GridPairs.ItemsSource = displayList;
                TxtFooterStatus.Text = $"✅ {result.Message}";

                if (result.Success)
                {
                    WpfMessageBox.Show(
                        $"{result.Message}\n\n• Tổng số lần thay thế: {result.TotalReplacements:N0}\n• Số ô đã cập nhật: {result.TotalCellsModified:N0}\n• Số Sheet đã xử lý: {result.SheetsModified:N0}",
                        "Thay Thế Hoàn Tất",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show(result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi thực thi:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtFooterStatus.Text = "❌ Đã xảy ra lỗi trong quá trình thay thế.";
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
