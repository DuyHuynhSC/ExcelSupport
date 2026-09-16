using System;
using ExcelDna.Integration;
using ExcelSupport.Host;
using ExcelSupport.Services;
using ExcelSupport.Views;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport
{
    /// <summary>
    /// Các macro Excel-DNA và Command phím tắt toàn cục cho Add-in
    /// </summary>
    public static class OracleCommands
    {
        [ExcelCommand(ShortCut = "^+Q", Name = "OracleQuickQueryCommand")]
        public static void OpenOracleQuickQuery()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                OracleQuickQueryDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
            });
        }
    }

    public static class DesignPageCounterCommands
    {
        [ExcelCommand(ShortCut = "^+H", Name = "ApplyDesignHighlightSelectionCommand")]
        public static void HighlightDesignSelection()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var app = (ExcelApp)ExcelDnaUtil.Application;
                    DesignPageCounterService.HighlightSelection(app);
                }
                catch { }
            });
        }

        [ExcelCommand(ShortCut = "^+%H", Name = "ClearDesignHighlightSelectionCommand")]
        public static void ClearDesignHighlightSelection()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var app = (ExcelApp)ExcelDnaUtil.Application;
                    DesignPageCounterService.ClearHighlightSelection(app);
                }
                catch { }
            });
        }
    }

    public static class TableExportCommands
    {
        [ExcelCommand(ShortCut = "^+M", Name = "ExportMarkdownTableCommand")]
        public static void ExportMarkdownTable()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var app = (ExcelApp)ExcelDnaUtil.Application;
                    TableExportService.QuickCopySelectionToMarkdown(app);
                }
                catch { }
            });
        }
    }

    public static class TaskPaneCommands
    {
        [ExcelCommand(ShortCut = "^+W", Name = "ToggleTaskPaneCommand")]
        public static void ToggleTaskPane()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    TaskPaneRegistry.ToggleTaskPaneAuto();
                }
                catch { }
            });
        }
    }

    public static class AiTranslationCommands
    {
        [ExcelCommand(ShortCut = "{F3}", Name = "OpenAiTranslateDialogCommand")]
        public static void OpenAiTranslateDialog()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    AiTranslateDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                catch { }
            });
        }

        [ExcelCommand(ShortCut = "^+T", Name = "AiQuickTranslateCommand")]
        public static void OpenAiQuickTranslatePopup()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    AiQuickTranslatePopup.ShowPopup(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                catch { }
            });
        }

        [ExcelCommand(ShortCut = "^%T", Name = "AiQuickTranslateAltCommand")]
        public static void OpenAiQuickTranslatePopupAlt()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    AiQuickTranslatePopup.ShowPopup(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                catch { }
            });
        }
    }

    public static class CommandPaletteCommands
    {
        [ExcelCommand(ShortCut = "^+P", Name = "QuickCommandPaletteCommand")]
        public static void OpenQuickCommandPalette()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    QuickCommandPaletteDialog.ShowWindow(AddInEvents.MainViewModel?.IsDarkTheme ?? false);
                }
                catch { }
            });
        }
    }

    public static class SpecLauncherCommands
    {
        [ExcelCommand(ShortCut = "^+D", Name = "OpenDetailedDesignCommand")]
        public static void OpenDetailedDesign()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                var result = ProjectDocumentLauncherService.LaunchFromSelection(app, Models.SpecDocumentType.DetailedDesign);
                if (!result.Success && !string.IsNullOrEmpty(result.Message))
                {
                    System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        [ExcelCommand(ShortCut = "^+B", Name = "OpenBasicDesignCommand")]
        public static void OpenBasicDesign()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                var result = ProjectDocumentLauncherService.LaunchFromSelection(app, Models.SpecDocumentType.BasicDesign);
                if (!result.Success && !string.IsNullOrEmpty(result.Message))
                {
                    System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        [ExcelCommand(ShortCut = "^+J", Name = "OpenTestSpecCommand")]
        public static void OpenTestSpec()
        {
            try
            {
                var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDnaUtil.Application;
                var result = ProjectDocumentLauncherService.LaunchFromSelection(app, Models.SpecDocumentType.TestSpec);
                if (!result.Success && !string.IsNullOrEmpty(result.Message))
                {
                    System.Windows.MessageBox.Show(result.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Thông Báo", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }
    }
}
