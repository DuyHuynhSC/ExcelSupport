using System;
using System.Windows.Forms;
using System.Windows.Input;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.ViewModels
{
    public class TableExportViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private bool _firstRowAsHeader = true;
        private bool _alignNumbersRight = true;
        private bool _compactFormat = false;
        private bool _includeHtmlStyles = true;
        private bool _convertLineBreaksToBr = true;

        private string _markdownText = string.Empty;
        private string _htmlText = string.Empty;
        private string _statusMessage = string.Empty;
        private int _selectedTabIndex = 0; // 0 = Markdown, 1 = HTML

        public bool FirstRowAsHeader
        {
            get => _firstRowAsHeader;
            set { if (SetProperty(ref _firstRowAsHeader, value)) RefreshExportContent(); }
        }

        public bool AlignNumbersRight
        {
            get => _alignNumbersRight;
            set { if (SetProperty(ref _alignNumbersRight, value)) RefreshExportContent(); }
        }

        public bool CompactFormat
        {
            get => _compactFormat;
            set { if (SetProperty(ref _compactFormat, value)) RefreshExportContent(); }
        }

        public bool IncludeHtmlStyles
        {
            get => _includeHtmlStyles;
            set { if (SetProperty(ref _includeHtmlStyles, value)) RefreshExportContent(); }
        }

        public bool ConvertLineBreaksToBr
        {
            get => _convertLineBreaksToBr;
            set { if (SetProperty(ref _convertLineBreaksToBr, value)) RefreshExportContent(); }
        }

        public string MarkdownText
        {
            get => _markdownText;
            set => SetProperty(ref _markdownText, value);
        }

        public string HtmlText
        {
            get => _htmlText;
            set => SetProperty(ref _htmlText, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ICommand CopyMarkdownCommand { get; }
        public ICommand CopyHtmlCommand { get; }
        public ICommand RefreshCommand { get; }

        public TableExportViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            CopyMarkdownCommand = new RelayCommand(_ =>
            {
                if (!string.IsNullOrEmpty(MarkdownText))
                {
                    Clipboard.SetText(MarkdownText);
                    StatusMessage = "📋 Đã sao chép bảng Markdown vào Clipboard!";
                }
            });

            CopyHtmlCommand = new RelayCommand(_ =>
            {
                if (!string.IsNullOrEmpty(HtmlText))
                {
                    Clipboard.SetText(HtmlText);
                    StatusMessage = "📋 Đã sao chép bảng HTML vào Clipboard!";
                }
            });

            RefreshCommand = new RelayCommand(_ => RefreshExportContent());

            RefreshExportContent();
        }

        public void RefreshExportContent()
        {
            try
            {
                dynamic sel = _excelApp.Selection;
                if (sel is not Range rng)
                {
                    MarkdownText = "Vui lòng chọn một vùng ô trên bảng tính Excel rồi bấm [Làm Mới].";
                    HtmlText = "<!-- Chưa chọn vùng ô trên Excel -->";
                    StatusMessage = "Chưa chọn vùng ô trên Excel.";
                    return;
                }

                var options = new TableExportOptions
                {
                    FirstRowAsHeader = FirstRowAsHeader,
                    AlignNumbersRight = AlignNumbersRight,
                    CompactFormat = CompactFormat,
                    IncludeHtmlStyles = IncludeHtmlStyles,
                    ConvertLineBreaksToBr = ConvertLineBreaksToBr
                };

                MarkdownText = TableExportService.RangeToMarkdown(rng, options);
                HtmlText = TableExportService.RangeToHtml(rng, options);
                StatusMessage = $"Đã xuất dữ liệu từ vùng {rng.Address} ({rng.Rows.Count} dòng x {rng.Columns.Count} cột).";
            }
            catch (Exception ex)
            {
                MarkdownText = $"Lỗi: {ex.Message}";
                HtmlText = $"<!-- Lỗi: {ex.Message} -->";
                StatusMessage = "Có lỗi xảy ra khi đọc vùng chọn Excel.";
            }
        }
    }
}
