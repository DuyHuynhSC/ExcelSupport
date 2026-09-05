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

        [ExcelCommand(ShortCut = "^+!H", Name = "ClearDesignHighlightSelectionCommand")]
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
}
