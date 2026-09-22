using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Views;
using Newtonsoft.Json;
using ExcelSupport.Ribbon;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using VmSortOrder = ExcelSupport.ViewModels.SortOrder;

namespace ExcelSupport.Services
{
    public static class CommandPaletteService
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelSupport"
        );
        private static readonly string RecentFilePath = Path.Combine(ConfigDirectory, "command_palette_recent.json");
        private static readonly object _recentLock = new object();
        private static List<RecentCommandEntry>? _cachedRecent;

        /// <summary>
        /// Tạo và lấy danh sách đầy đủ tất cả lệnh trong hệ thống
        /// </summary>
        public static List<PaletteCommandItem> GetAllCommands()
        {
            var isDark = AddInEvents.MainViewModel?.IsDarkTheme ?? false;

            return new List<PaletteCommandItem>
            {
                // ==================== 1. ĐIỀU HƯỚNG & HỆ THỐNG (grpNavExplorer) ====================
                new PaletteCommandItem
                {
                    Id = "toggle_command_palette",
                    TitleKey = "btnCommandPalette",
                    DescriptionKey = "btnCommandPalette_SuperTip",
                    CategoryKey = "grpNavExplorer",
                    IconEmoji = "⚡",
                    ShortcutText = "Ctrl + Shift + P",
                    Keywords = new List<string> { "command palette", "menu", "quick", "tim kiem", "search", "raycast", "spotlight", "lenh", "phim tat", "actions" },
                    Action = () => QuickCommandPaletteDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "toggle_taskpane",
                    TitleKey = "btnToggleTaskPane",
                    DescriptionKey = "btnToggleTaskPane_SuperTip",
                    CategoryKey = "grpNavExplorer",
                    IconEmoji = "📑",
                    ShortcutText = "Ctrl + Shift + W",
                    Keywords = new List<string> { "taskpane", "navigator", "cay sheet", "sheet tree", "side bar", "thanh dieu huong", "dieu huong", "explorer" },
                    Action = () => TaskPaneRegistry.ToggleTaskPaneAuto()
                },
                new PaletteCommandItem
                {
                    Id = "refresh_tree",
                    TitleKey = "btnRefreshTree",
                    DescriptionKey = "btnRefreshTree_SuperTip",
                    CategoryKey = "grpNavExplorer",
                    IconEmoji = "🔄",
                    Keywords = new List<string> { "refresh", "lam moi", "quet lai", "reload", "cap nhat danh sach" },
                    Action = () => AddInEvents.Instance?.RefreshWorkbookTreePublic()
                },
                new PaletteCommandItem
                {
                    Id = "toggle_quick_action_bar",
                    TitleKey = "btnToggleQuickActionBar",
                    DescriptionKey = "btnToggleQuickActionBar_SuperTip",
                    CategoryKey = "grpNavExplorer",
                    IconEmoji = "🪄",
                    Keywords = new List<string> { "quick action bar", "floating bar", "thanh tac vu noi", "con tro", "cursor", "mini toolbar", "quick action", "selection bar" },
                    Action = () =>
                    {
                        QuickActionBarService.ToggleEnabled();
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        string stateMsg = AppSettings.IsQuickActionBarEnabled ? "BẬT" : "TẮT";
                        try { if (app != null) app.StatusBar = $"✨ ExcelSupport: Đã {stateMsg} Thanh tác vụ nổi (Mini Floating Bar)!"; } catch { }
                    }
                },

                // ==================== 2. THAO TÁC NHANH (grpQuickTools) ====================
                new PaletteCommandItem
                {
                    Id = "create_toc",
                    TitleKey = "btnCreateTOC",
                    DescriptionKey = "btnCreateTOC_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "📑",
                    Keywords = new List<string> { "toc", "muc luc", "table of contents", "index", "tao trang muc luc", "hyperlink" },
                    Action = () => ExcelAsyncUtil.QueueAsMacro(() => AddInEvents.Instance?.CreateTableOfContents(null))
                },
                new PaletteCommandItem
                {
                    Id = "split_sheets",
                    TitleKey = "btnSplitSheets",
                    DescriptionKey = "btnSplitSheets_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "📤",
                    Keywords = new List<string> { "split", "tach sheet", "xuat sheet", "chia file", "export separate files" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
                        {
                            var dlg = new SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 0, AddInEvents.MainViewModel.IsDarkTheme);
                            dlg.ShowDialog();
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "merge_sheets",
                    TitleKey = "btnMergeSheets",
                    DescriptionKey = "btnMergeSheets_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "📥",
                    Keywords = new List<string> { "merge sheets", "gop sheet", "tong hop sheet", "consolidate" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
                        {
                            var dlg = new SheetToolsDialog(AddInEvents.MainViewModel.SelectedWorkbook, 1, AddInEvents.MainViewModel.IsDarkTheme);
                            dlg.ShowDialog();
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "sort_sheets_az",
                    TitleKey = "btnQuickSortAZ",
                    DescriptionKey = "btnQuickSortAZ_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "🔤",
                    Keywords = new List<string> { "sort az", "sap xep sheet az", "a den z", "tang dan", "alphabetical" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel != null)
                        {
                            AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Ascending;
                            AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Ascending;
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "sort_sheets_za",
                    TitleKey = "btnQuickSortZA",
                    DescriptionKey = "btnQuickSortZA_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "🔤",
                    Keywords = new List<string> { "sort za", "sap xep sheet za", "z den a", "giam dan", "reverse" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel != null)
                        {
                            AddInEvents.MainViewModel.WorkbookSortOrder = VmSortOrder.Descending;
                            AddInEvents.MainViewModel.SheetSortOrder = VmSortOrder.Descending;
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "close_current_workbook",
                    TitleKey = "btnCloseCurrentWb",
                    DescriptionKey = "btnCloseCurrentWb_SuperTip",
                    CategoryKey = "grpQuickTools",
                    IconEmoji = "❌",
                    Keywords = new List<string> { "close", "dong workbook", "dong file", "exit", "close file" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel?.SelectedWorkbook != null)
                        {
                            AddInEvents.MainViewModel.CloseWorkbookCommand.Execute(AddInEvents.MainViewModel.SelectedWorkbook.WorkbookName);
                        }
                    }
                },

                // ==================== 3. KIỂM TRA & ĐỐI SOÁT (grpAuditTools) ====================
                new PaletteCommandItem
                {
                    Id = "compare_workbooks",
                    TitleKey = "btnCompareWorkbooks",
                    DescriptionKey = "btnCompareWorkbooks_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "⚖️",
                    Keywords = new List<string> { "compare", "so sanh", "diff", "doi chieu", "kiem tra lech", "so sanh 2 file excel" },
                    Action = () => WorkbookCompareDialog.ShowWindow(AddInEvents.MainViewModel?.SelectedWorkbook?.WorkbookName, AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "check_vietnamese",
                    TitleKey = "btnCheckVietnamese",
                    DescriptionKey = "btnCheckVietnamese_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "🇻🇳",
                    Keywords = new List<string> { "vietnamese", "tieng viet", "kiem tra tieng viet", "tieng viet co dau", "accent check", "audit" },
                    Action = () => VietnameseCheckDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "external_links_manager",
                    TitleKey = "btnExternalLinks",
                    DescriptionKey = "btnExternalLinks_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "🔗",
                    Keywords = new List<string> { "external links", "break link", "lien ket ngoai", "sua loi link", "ngat ket noi file" },
                    Action = () => ExternalLinksManagerDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "oracle_table_compare",
                    TitleKey = "btnOracleTableCompare",
                    DescriptionKey = "btnOracleTableCompare_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "🗄️",
                    Keywords = new List<string> { "oracle compare", "so sanh oracle", "database diff", "so sanh bang", "schema compare", "oracle" },
                    Action = () => OracleTableCompareDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "oracle_quick_query",
                    TitleKey = "btnOracleQuickQuery",
                    DescriptionKey = "btnOracleQuickQuery_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "⚡",
                    ShortcutText = "Ctrl + Shift + Q",
                    Keywords = new List<string> { "oracle query", "truy van oracle", "sql", "select", "database query", "oracle" },
                    Action = () => OracleQuickQueryDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "sql_script_generator",
                    TitleKey = "btnSqlScriptGenerator",
                    DescriptionKey = "btnSqlScriptGenerator_SuperTip",
                    CategoryKey = "grpAuditTools",
                    IconEmoji = "🗄️",
                    ShortcutText = "Ctrl + Shift + K",
                    Keywords = new List<string> { "sql", "range to sql", "insert", "merge", "upsert", "script", "database", "sinh sql", "xuat sql", "tao insert", "table", "cau lenh sql" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        SqlScriptGeneratorDialog.ShowWindow(app, AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                    }
                },

                // ==================== 4. XỬ LÝ DỮ LIỆU (grpDataTools) ====================
                new PaletteCommandItem
                {
                    Id = "advanced_filter",
                    TitleKey = "btnAdvancedFilter",
                    DescriptionKey = "btnAdvancedFilter_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "⚡",
                    Keywords = new List<string> { "filter", "loc nang cao", "advanced filter", "loc du lieu", "smart filter", "multi condition" },
                    Action = () => AdvancedFilterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "data_cleaner",
                    TitleKey = "btnDataCleaner",
                    DescriptionKey = "btnDataCleaner_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "🧹",
                    Keywords = new List<string> { "clean", "don dep du lieu", "chuan hoa", "trim", "xoa khoang trang", "upper", "lower", "cleaner" },
                    Action = () => DataCleaningWizardDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "filtered_copy_paste_wizard",
                    TitleKey = "btnFilteredCopyPasteWizard",
                    DescriptionKey = "btnFilteredCopyPasteWizard_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📋",
                    Keywords = new List<string> { "filtered paste", "dan co loc", "copy paste visible", "tro ly dan" },
                    Action = () => FilteredCopyPasteDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "copy_visible_only",
                    TitleKey = "btnCopyVisibleOnly",
                    DescriptionKey = "btnCopyVisibleOnly_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📑",
                    Keywords = new List<string> { "copy visible", "sao chep o hien thi", "chi copy o hien", "copy non hidden" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        var result = FilteredCopyPasteService.CopyVisibleCells(app);
                        if (!result.Success)
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "paste_to_visible_only",
                    TitleKey = "btnPasteToVisibleOnly",
                    DescriptionKey = "btnPasteToVisibleOnly_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📋",
                    Keywords = new List<string> { "paste visible", "dan vao o hien thi", "chi dan vao dong hien", "skip hidden" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        var result = FilteredCopyPasteService.PasteToVisibleCells(app);
                        if (result.Success)
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                        else
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "special_copy_wizard",
                    TitleKey = "btnSpecialCopyWizard",
                    DescriptionKey = "btnSpecialCopyWizard_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "⚡",
                    Keywords = new List<string> { "special copy", "noi chuoi", "concatenate", "csv row", "join cells", "copy delimiter", "phay", "cham phay", "pipe" },
                    Action = () => SpecialCopyDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "special_copy_comma",
                    TitleKey = "btnSpecialCopyComma",
                    DescriptionKey = "btnSpecialCopyComma_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "⚡",
                    Keywords = new List<string> { "special copy comma", "noi dau phay", "join comma", "copy comma" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        var result = SpecialCopyService.QuickCopy(app, SpecialCopyDelimiter.Comma);
                        if (result.Success)
                            System.Windows.MessageBox.Show($"{result.Message}\n\n👉 Nhấn Ctrl + V để dán vào ô tính hoặc ứng dụng khác!", LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                        else
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "special_copy_pipe",
                    TitleKey = "btnSpecialCopyPipe",
                    DescriptionKey = "btnSpecialCopyPipe_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "⚡",
                    Keywords = new List<string> { "special copy pipe", "noi dau pipe", "join pipe", "gach dung" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        var result = SpecialCopyService.QuickCopy(app, SpecialCopyDelimiter.Pipe);
                        if (result.Success)
                            System.Windows.MessageBox.Show($"{result.Message}\n\n👉 Nhấn Ctrl + V để dán vào ô tính hoặc ứng dụng khác!", LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                        else
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "special_paste",
                    TitleKey = "btnSpecialPaste",
                    DescriptionKey = "btnSpecialPaste_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📥",
                    Keywords = new List<string> { "special paste", "dan chuoi noi", "paste joined" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        var result = SpecialCopyService.PasteSpecialCopy(app);
                        if (result.Success)
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                        else
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("btnCommandPalette"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "duplicate_finder",
                    TitleKey = "btnDuplicateFinder",
                    DescriptionKey = "btnDuplicateFinder_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📑",
                    Keywords = new List<string> { "duplicate", "tim trung", "dong trung", "lap", "loc trung", "distinct", "unique" },
                    Action = () => DuplicateFinderDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "batch_blank_cleaner",
                    TitleKey = "btnBatchBlankCleaner",
                    DescriptionKey = "btnBatchBlankCleaner_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "⛔",
                    Keywords = new List<string> { "blank", "xoa dong trong", "xoa cot trong", "empty rows", "delete empty" },
                    Action = () => BatchCleanerAndMergeDialog.ShowWindow(0, AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "batch_find_replace",
                    TitleKey = "btnBatchFindReplace",
                    DescriptionKey = "btnBatchFindReplace_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "🔍",
                    Keywords = new List<string> { "find replace", "tim va thay the hang loat", "thay the nhieu tu", "batch replace", "dictionary replace" },
                    Action = () => BatchFindReplaceDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "visual_table_merge",
                    TitleKey = "btnVisualTableMerge",
                    DescriptionKey = "btnVisualTableMerge_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "📊",
                    Keywords = new List<string> { "vlookup", "table merge", "gop bang", "join", "xlookup", "merge 2 bang" },
                    Action = () => VisualTableMergeDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "fuzzy_duplicate",
                    TitleKey = "btnFuzzyDuplicate",
                    DescriptionKey = "btnFuzzyDuplicate_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "🔄",
                    Keywords = new List<string> { "fuzzy duplicate", "trung lap gan dung", "trung lap ao", "tuong tu", "similarity" },
                    Action = () => FuzzyDuplicateDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "safe_merge_consolidate",
                    TitleKey = "btnSafeMergeConsolidate",
                    DescriptionKey = "btnSafeMergeConsolidate_SuperTip",
                    CategoryKey = "grpDataTools",
                    IconEmoji = "🔀",
                    Keywords = new List<string> { "safe merge", "gop du lieu an toan", "merge preserve data", "consolidate" },
                    Action = () => BatchCleanerAndMergeDialog.ShowWindow(1, AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },

                // ==================== 5. QUẢN TRỊ TẬP TIN & THIẾT KẾ (grpFileTools) ====================
                new PaletteCommandItem
                {
                    Id = "batch_file_converter",
                    TitleKey = "btnBatchFileConverter",
                    DescriptionKey = "btnBatchFileConverter_SuperTip",
                    CategoryKey = "grpFileTools",
                    IconEmoji = "🔄",
                    Keywords = new List<string> { "convert", "chuyen doi file", "xls to xlsx", "csv", "pdf", "batch convert" },
                    Action = () => BatchFileConverterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "design_page_counter",
                    TitleKey = "btnDesignPageCounter",
                    DescriptionKey = "btnDesignPageCounter_SuperTip",
                    CategoryKey = "grpFileTools",
                    IconEmoji = "📑",
                    Keywords = new List<string> { "dem trang", "design page counter", "thiet ke 2.0", "page count", "trang in", "print area" },
                    Action = () => DesignPageCounterDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "highlight_design_selection",
                    TitleKey = "HighlightDesignSelectionTitle",
                    DescriptionKey = "HighlightDesignSelectionTip",
                    CategoryKey = "grpFileTools",
                    IconEmoji = "🖍️",
                    ShortcutText = "Ctrl + Shift + H",
                    Keywords = new List<string> { "highlight design", "to mau thiet ke", "highlight selection", "design border" },
                    Action = () => ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        var app = (ExcelApp)ExcelDnaUtil.Application;
                        DesignPageCounterService.HighlightSelection(app);
                    })
                },
                new PaletteCommandItem
                {
                    Id = "clear_design_highlight",
                    TitleKey = "ClearDesignHighlightTitle",
                    DescriptionKey = "ClearDesignHighlightTip",
                    CategoryKey = "grpFileTools",
                    IconEmoji = "🧹",
                    ShortcutText = "Ctrl + Shift + Alt + H",
                    Keywords = new List<string> { "clear highlight", "xoa to mau thiet ke", "remove highlight", "unhighlight" },
                    Action = () => ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        var app = (ExcelApp)ExcelDnaUtil.Application;
                        DesignPageCounterService.ClearHighlightSelection(app);
                    })
                },

                // ==================== 6. TIỆN ÍCH IT & KHÁCH HÀNG NHẬT (grpJapanTools) ====================
                new PaletteCommandItem
                {
                    Id = "launch_detailed_design",
                    TitleKey = "btnLaunchDetailedDesign",
                    DescriptionKey = "btnLaunchDetailedDesign_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "📄",
                    ShortcutText = "Ctrl + Shift + D",
                    Keywords = new List<string> { "tkct", "detailed design", "thiet ke chi tiet", "spec", "screen id", "table id", "launcher", "mo tkct" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var result = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.DetailedDesign);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "launch_basic_design",
                    TitleKey = "btnLaunchBasicDesign",
                    DescriptionKey = "btnLaunchBasicDesign_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "📐",
                    ShortcutText = "Ctrl + Shift + B",
                    Keywords = new List<string> { "tkcb", "basic design", "thiet ke co ban", "spec", "screen id", "table id", "launcher", "mo tkcb" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var result = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.BasicDesign);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "launch_test_spec",
                    TitleKey = "btnLaunchTestSpec",
                    DescriptionKey = "btnLaunchTestSpec_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "🧪",
                    ShortcutText = "Ctrl + Shift + J",
                    Keywords = new List<string> { "test spec", "chi thi test", "test case", "chi thi", "kiem thu", "launcher", "mo test" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var result = ProjectDocumentLauncherService.LaunchFromSelection(app, SpecDocumentType.TestSpec);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "project_profile_settings",
                    TitleKey = "btnProjectProfileSettings",
                    DescriptionKey = "btnProjectProfileSettings_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "⚙️",
                    Keywords = new List<string> { "project profile", "du an", "cau hinh du an", "profile setting", "spec folder", "thu muc thiet ke" },
                    Action = () => ProjectProfileSettingsDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },

                // ==================== JAPANESE PRESET FORMATTER ====================
                new PaletteCommandItem
                {
                    Id = "jp_format_default",
                    TitleKey = "btnApplyDefaultJapanesePreset",
                    DescriptionKey = "btnApplyDefaultJapanesePreset_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "🎨",
                    Keywords = new List<string> { "format nhat", "japanese format", "chuan hoa thiet ke", "dinh dang tkct", "preset nhat", "style tkct", "japanese style", "meiryo" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetActivePreset();
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_header",
                    TitleKey = "btnApplyPresetTableHeader",
                    DescriptionKey = "btnApplyPresetTableHeader_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "📑",
                    Keywords = new List<string> { "header tkct", "table header", "tieu de bang", "navy header", "dinh dang header", "cot tkct" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetPresetById("builtin_table_header");
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_must",
                    TitleKey = "btnApplyPresetMust",
                    DescriptionKey = "btnApplyPresetMust_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "🔴",
                    Keywords = new List<string> { "must", "bat buoc", "required", "truong bat buoc", "not null", "peach red" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetPresetById("builtin_required_must");
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_code",
                    TitleKey = "btnApplyPresetCode",
                    DescriptionKey = "btnApplyPresetCode_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "💻",
                    Keywords = new List<string> { "code", "physical name", "ten vat ly", "consolas", "ma chuong trinh", "table id", "field id" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetPresetById("builtin_code_physical");
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_note",
                    TitleKey = "btnApplyPresetNote",
                    DescriptionKey = "btnApplyPresetNote_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "💡",
                    Keywords = new List<string> { "note", "ghi chu", "canh bao", "warning", "luu y", "amber", "yellow" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetPresetById("builtin_note_warning");
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_standard",
                    TitleKey = "btnApplyPresetStandard",
                    DescriptionKey = "btnApplyPresetStandard_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "📄",
                    Keywords = new List<string> { "standard data", "du lieu chuan", "dong noi dung", "meiryo 9.5", "clear format to standard" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        var preset = JapanesePresetFormatManager.GetPresetById("builtin_standard_data");
                        var result = JapanesePresetFormatService.ApplyPresetToSelection(app, preset);
                        if (!result.Success && !string.IsNullOrEmpty(result.Message))
                        {
                            System.Windows.MessageBox.Show(result.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "jp_format_manager",
                    TitleKey = "btnOpenJapanesePresetManager",
                    DescriptionKey = "btnOpenJapanesePresetManager_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "⚙️",
                    Keywords = new List<string> { "quan ly preset", "japanese preset manager", "cai dat dinh dang nhat", "format config", "presets dialog", "live preview" },
                    Action = () => JapanesePresetManagerDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "japanese_convert",
                    TitleKey = "btnJapaneseConvert",
                    DescriptionKey = "btnJapaneseConvert_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "🇯🇵",
                    Keywords = new List<string> { "zenkaku", "hankaku", "toan giac", "ban giac", "nhat ban", "japanese", "fullwidth", "halfwidth" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        JapaneseTextConverterDialog.ShowWindow(app);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "katakana_check",
                    TitleKey = "btnKatakanaCheck",
                    DescriptionKey = "btnKatakanaCheck_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "🈁",
                    Keywords = new List<string> { "katakana", "choon", "chuan hoa katakana", "tieng nhat", "validate katakana" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        KatakanaValidatorDialog.ShowWindow(app);
                    }
                },
                new PaletteCommandItem
                {
                    Id = "export_markdown",
                    TitleKey = "btnExportMarkdown",
                    DescriptionKey = "btnExportMarkdown_SuperTip",
                    CategoryKey = "grpJapanTools",
                    IconEmoji = "📝",
                    ShortcutText = "Ctrl + Shift + M",
                    Keywords = new List<string> { "markdown", "xuat bang markdown", "jira table", "html table", "export markdown", "copy markdown" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                        TableExportDialog.ShowWindow(app);
                    }
                },

                // ==================== 7. THƯỚC NGẮM & HIỂN THỊ (grpViewTools) ====================
                new PaletteCommandItem
                {
                    Id = "toggle_grid_ruler",
                    TitleKey = "btnToggleGridRuler",
                    DescriptionKey = "btnToggleGridRuler_SuperTip",
                    CategoryKey = "grpViewTools",
                    IconEmoji = "📐",
                    Keywords = new List<string> { "ruler", "thuoc ngam", "soi dong", "crosshair", "grid ruler", "highlight dong cot" },
                    Action = () =>
                    {
                        var app = AddInEvents.Instance?.ExcelAppInstance;
                        GridRulerService.Toggle(app);
                        RibbonController.Instance?.InvalidateControl("btnToggleGridRuler");
                    }
                },
                new PaletteCommandItem
                {
                    Id = "toggle_ruler_hud",
                    TitleKey = "btnToggleHud",
                    DescriptionKey = "btnToggleHud_Tip",
                    CategoryKey = "grpViewTools",
                    IconEmoji = "📊",
                    Keywords = new List<string> { "hud", "bang thong ke noi", "floating hud", "dynamic hud", "thong ke o" },
                    Action = () => RulerHudWindow.ForceOpenHud(AddInEvents.MainViewModel?.IsDarkTheme ?? true)
                },

                // ==================== 8. TRỢ LÝ AI & NĂNG SUẤT (grpAiTools) ====================
                new PaletteCommandItem
                {
                    Id = "ai_quick_translate",
                    TitleKey = "btnAiQuickTranslate",
                    DescriptionKey = "btnAiQuickTranslate_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "🌐",
                    ShortcutText = "Ctrl + Shift + T / F3",
                    Keywords = new List<string> { "dich nhanh", "quick translate", "dich ai", "dich nhat viet", "translate cell", "f3", "ai" },
                    Action = () => AiQuickTranslatePopup.ShowPopup(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "ai_translate",
                    TitleKey = "btnAiTranslate",
                    DescriptionKey = "btnAiTranslate_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "🌐",
                    Keywords = new List<string> { "ai translate", "hop thoai dich", "dich thuat chuyen sau", "translate dialog", "dich nhieu o" },
                    Action = () => AiTranslateDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "ai_formula",
                    TitleKey = "btnAiFormula",
                    DescriptionKey = "btnAiFormula_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "✨",
                    Keywords = new List<string> { "ai formula", "sinh cong thuc", "viet cong thuc ai", "formula copilot", "generate formula" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel != null)
                        {
                            TaskPaneRegistry.ToggleTaskPane(AddInEvents.MainViewModel, true);
                            AddInEvents.MainViewModel.SelectedTabIndex = 1;
                            if (AddInEvents.MainViewModel.AiAssistant != null)
                            {
                                AddInEvents.MainViewModel.AiAssistant.SelectedSubTab = 1;
                            }
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "ai_formula_doctor",
                    TitleKey = "btnAiFormulaDoctor",
                    DescriptionKey = "btnAiFormulaDoctor_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "🩺",
                    Keywords = new List<string> { "formula doctor", "bac si cong thuc", "sua loi cong thuc", "fix formula", "giai thich cong thuc" },
                    Action = () => AiFormulaDoctorDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "snapshot_rollback",
                    TitleKey = "btnSnapshotRollback",
                    DescriptionKey = "btnSnapshotRollback_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "📸",
                    Keywords = new List<string> { "snapshot", "sao luu", "khoi phuc", "rollback", "undo", "luu diem phuc hoi", "instant undo" },
                    Action = () => SheetSnapshotDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },
                new PaletteCommandItem
                {
                    Id = "user_manual",
                    TitleKey = "btnUserManual",
                    DescriptionKey = "btnUserManual_SuperTip",
                    CategoryKey = "grpAiTools",
                    IconEmoji = "📖",
                    Keywords = new List<string> { "user manual", "huong dan su dung", "help", "tro giup", "tai lieu", "documentation" },
                    Action = () => UserManualDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false)
                },

                // ==================== 9. CÀI ĐẶT & HỆ THỐNG (grpSettings) ====================
                new PaletteCommandItem
                {
                    Id = "customize_ribbon",
                    TitleKey = "btnCustomizeRibbon",
                    DescriptionKey = "btnCustomizeRibbon_SuperTip",
                    CategoryKey = "grpSettings",
                    IconEmoji = "⚙️",
                    Keywords = new List<string> { "settings", "cai dat", "tuy chinh ribbon", "customize", "an hien nut", "giao dien" },
                    Action = () => RibbonCustomizeDialog.ShowWindow(0)
                },
                new PaletteCommandItem
                {
                    Id = "ai_settings",
                    TitleKey = "AiSettings_Title",
                    DescriptionKey = "AiSettings_Desc",
                    CategoryKey = "grpSettings",
                    IconEmoji = "🤖",
                    Keywords = new List<string> { "ai settings", "api key", "cau hinh ai", "openai", "gemini", "claude", "deepseek" },
                    Action = () => RibbonCustomizeDialog.ShowWindow(1)
                },
                new PaletteCommandItem
                {
                    Id = "toggle_theme",
                    TitleKey = "ToggleTheme_Title",
                    DescriptionKey = "ToggleTheme_Desc",
                    CategoryKey = "grpSettings",
                    IconEmoji = "🌗",
                    Keywords = new List<string> { "theme", "dark mode", "light mode", "giao dien toi", "giao dien sang", "mau sac" },
                    Action = () =>
                    {
                        if (AddInEvents.MainViewModel != null)
                        {
                            AddInEvents.MainViewModel.IsDarkTheme = !AddInEvents.MainViewModel.IsDarkTheme;
                        }
                    }
                },
                new PaletteCommandItem
                {
                    Id = "lang_vi",
                    TitleKey = "LangVi_Title",
                    DescriptionKey = "LangVi_Desc",
                    CategoryKey = "grpSettings",
                    IconEmoji = "🇻🇳",
                    Keywords = new List<string> { "tieng viet", "vietnamese", "ngon ngu", "language vi" },
                    Action = () =>
                    {
                        LocalizationService.CurrentLanguage = AppLanguage.Vietnamese;
                        RibbonController.Instance?.InvalidateRibbon();
                    }
                },
                new PaletteCommandItem
                {
                    Id = "lang_en",
                    TitleKey = "LangEn_Title",
                    DescriptionKey = "LangEn_Desc",
                    CategoryKey = "grpSettings",
                    IconEmoji = "🇬🇧",
                    Keywords = new List<string> { "english", "tieng anh", "language en" },
                    Action = () =>
                    {
                        LocalizationService.CurrentLanguage = AppLanguage.English;
                        RibbonController.Instance?.InvalidateRibbon();
                    }
                },
                new PaletteCommandItem
                {
                    Id = "lang_ja",
                    TitleKey = "LangJa_Title",
                    DescriptionKey = "LangJa_Desc",
                    CategoryKey = "grpSettings",
                    IconEmoji = "🇯🇵",
                    Keywords = new List<string> { "japanese", "tieng nhat", "nihongo", "language ja" },
                    Action = () =>
                    {
                        LocalizationService.CurrentLanguage = AppLanguage.Japanese;
                        RibbonController.Instance?.InvalidateRibbon();
                    }
                }
            };
        }

        /// <summary>
        /// Tìm kiếm và xếp hạng các lệnh theo chuỗi truy vấn (Fuzzy Search)
        /// </summary>
        public static List<PaletteCommandItem> Search(string? query)
        {
            var all = GetAllCommands();
            var recentIds = GetRecentCommandIds();
            var recentSet = new HashSet<string>(recentIds, StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(query))
            {
                // Khi không gõ gì: Hiện các lệnh Recent lên đầu, sau đó là toàn bộ lệnh
                var recentList = new List<PaletteCommandItem>();
                foreach (var id in recentIds)
                {
                    var found = all.FirstOrDefault(c => string.Equals(c.Id, id, StringComparison.OrdinalIgnoreCase));
                    if (found != null)
                    {
                        found.IsRecent = true;
                        recentList.Add(found);
                    }
                }

                var remainder = all
                    .Where(c => !recentSet.Contains(c.Id))
                    .OrderBy(c => c.Category)
                    .ThenBy(c => c.Title)
                    .ToList();

                return recentList.Concat(remainder).ToList();
            }

            string rawQuery = query!.Trim().ToLowerInvariant();
            string unaccentedQuery = RemoveDiacritics(rawQuery);
            var queryTokens = unaccentedQuery.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            var scoredList = new List<PaletteCommandItem>();

            foreach (var cmd in all)
            {
                int score = CalculateMatchScore(cmd, rawQuery, unaccentedQuery, queryTokens);
                if (recentSet.Contains(cmd.Id))
                {
                    cmd.IsRecent = true;
                    score += 40; // Điểm ưu tiên nhẹ cho lệnh vừa dùng
                }

                if (score > 0)
                {
                    cmd.MatchScore = score;
                    scoredList.Add(cmd);
                }
            }

            return scoredList
                .OrderByDescending(c => c.MatchScore)
                .ThenBy(c => c.Title)
                .ToList();
        }

        private static int CalculateMatchScore(PaletteCommandItem cmd, string rawQuery, string unaccentedQuery, string[] queryTokens)
        {
            int score = 0;

            string rawTitle = cmd.Title.ToLowerInvariant();
            string unaccentedTitle = RemoveDiacritics(rawTitle);

            string rawDesc = cmd.Description.ToLowerInvariant();
            string unaccentedDesc = RemoveDiacritics(rawDesc);

            string rawCategory = cmd.Category.ToLowerInvariant();
            string unaccentedCategory = RemoveDiacritics(rawCategory);

            string rawShortcut = cmd.ShortcutText.ToLowerInvariant().Replace(" ", "");
            string queryNoSpace = rawQuery.Replace(" ", "");

            // 1. Phím tắt trùng khớp
            if (!string.IsNullOrEmpty(rawShortcut) && rawShortcut.Contains(queryNoSpace))
            {
                score += 800;
            }

            // 2. Khớp tuyệt đối hoặc bắt đầu bằng query trong Title
            if (rawTitle == rawQuery || unaccentedTitle == unaccentedQuery)
            {
                score += 1000;
            }
            else if (rawTitle.StartsWith(rawQuery) || unaccentedTitle.StartsWith(unaccentedQuery))
            {
                score += 700;
            }
            else if (rawTitle.Contains(rawQuery) || unaccentedTitle.Contains(unaccentedQuery))
            {
                score += 500;
            }

            // 3. Khớp Acronym viết tắt trong Title (ví dụ "af" -> Advanced Filter, "df" -> Duplicate Finder)
            string titleAcronym = GetAcronym(unaccentedTitle);
            if (!string.IsNullOrEmpty(titleAcronym) && titleAcronym.StartsWith(unaccentedQuery))
            {
                score += 450;
            }

            // 4. Khớp Keywords
            foreach (var kw in cmd.Keywords)
            {
                string normKw = RemoveDiacritics(kw.ToLowerInvariant());
                if (normKw == unaccentedQuery)
                {
                    score += 600;
                }
                else if (normKw.StartsWith(unaccentedQuery))
                {
                    score += 400;
                }
                else if (normKw.Contains(unaccentedQuery))
                {
                    score += 300;
                }
            }

            // 5. Khớp mô tả (Description) hoặc danh mục (Category)
            if (unaccentedDesc.Contains(unaccentedQuery))
            {
                score += 150;
            }
            if (unaccentedCategory.Contains(unaccentedQuery))
            {
                score += 120;
            }

            // 6. Fuzzy Subsequence (tất cả các ký tự xuất hiện theo thứ tự trong Title)
            if (score == 0 && IsSubsequence(unaccentedQuery, unaccentedTitle, out int compactness))
            {
                score += Math.Max(50, 200 - compactness);
            }

            // 7. Khớp đa từ khóa (nếu gõ nhiều từ như "dich nhanh")
            if (queryTokens.Length > 1)
            {
                bool allTokensMatch = queryTokens.All(tok =>
                    unaccentedTitle.Contains(tok) ||
                    unaccentedDesc.Contains(tok) ||
                    cmd.Keywords.Any(k => RemoveDiacritics(k.ToLowerInvariant()).Contains(tok))
                );

                if (allTokensMatch)
                {
                    score += 350;
                }
            }

            return score;
        }

        private static string GetAcronym(string text)
        {
            var parts = text.Split(new[] { ' ', '-', '/', '_' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            foreach (var p in parts)
            {
                if (p.Length > 0 && char.IsLetterOrDigit(p[0]))
                {
                    sb.Append(p[0]);
                }
            }
            return sb.ToString().ToLowerInvariant();
        }

        private static bool IsSubsequence(string sub, string target, out int distance)
        {
            distance = int.MaxValue;
            if (string.IsNullOrEmpty(sub) || string.IsNullOrEmpty(target) || sub.Length > target.Length)
                return false;

            int subIndex = 0;
            int firstMatch = -1;
            int lastMatch = -1;

            for (int i = 0; i < target.Length; i++)
            {
                if (target[i] == sub[subIndex])
                {
                    if (firstMatch == -1) firstMatch = i;
                    lastMatch = i;
                    subIndex++;
                    if (subIndex == sub.Length)
                    {
                        distance = lastMatch - firstMatch;
                        return true;
                    }
                }
            }

            return false;
        }

        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char c in normalized)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    // Chuyển đ -> d, Đ -> D
                    if (c == 'đ') sb.Append('d');
                    else if (c == 'Đ') sb.Append('D');
                    else sb.Append(c);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        // ==================== RECENT COMMANDS PERSISTENCE ====================

        public static List<string> GetRecentCommandIds()
        {
            lock (_recentLock)
            {
                if (_cachedRecent != null)
                {
                    return _cachedRecent.Select(r => r.CommandId).ToList();
                }

                try
                {
                    if (File.Exists(RecentFilePath))
                    {
                        string json = File.ReadAllText(RecentFilePath, Encoding.UTF8);
                        var list = JsonConvert.DeserializeObject<List<RecentCommandEntry>>(json);
                        if (list != null)
                        {
                            _cachedRecent = list.OrderByDescending(r => r.LastUsedUtc).ToList();
                            return _cachedRecent.Select(r => r.CommandId).ToList();
                        }
                    }
                }
                catch { }

                _cachedRecent = new List<RecentCommandEntry>();
                return new List<string>();
            }
        }

        public static void RecordCommandExecution(string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId)) return;

            lock (_recentLock)
            {
                try
                {
                    GetRecentCommandIds(); // ensure loaded
                    _cachedRecent ??= new List<RecentCommandEntry>();

                    _cachedRecent.RemoveAll(r => string.Equals(r.CommandId, commandId, StringComparison.OrdinalIgnoreCase));
                    _cachedRecent.Insert(0, new RecentCommandEntry
                    {
                        CommandId = commandId,
                        LastUsedUtc = DateTime.UtcNow
                    });

                    // Giới hạn tối đa 15 lệnh gần nhất
                    if (_cachedRecent.Count > 15)
                    {
                        _cachedRecent = _cachedRecent.Take(15).ToList();
                    }

                    if (!Directory.Exists(ConfigDirectory))
                    {
                        Directory.CreateDirectory(ConfigDirectory);
                    }

                    string json = JsonConvert.SerializeObject(_cachedRecent, Formatting.Indented);
                    File.WriteAllText(RecentFilePath, json, Encoding.UTF8);
                }
                catch { }
            }
        }
    }
}
