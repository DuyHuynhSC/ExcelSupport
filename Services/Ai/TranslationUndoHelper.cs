using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Services
{
    public static class TranslationUndoHelper
    {
        public class CellBackupItem
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public object? OldValue { get; set; }
            public object? NewValue { get; set; }
        }

        public class UndoRecord
        {
            public string SheetName { get; set; } = string.Empty;
            public string? WorkbookName { get; set; }
            public List<CellBackupItem> Cells { get; set; } = new List<CellBackupItem>();
        }

        private static readonly Stack<UndoRecord> _undoStack = new Stack<UndoRecord>();
        private static readonly Stack<UndoRecord> _redoStack = new Stack<UndoRecord>();

        public const string UndoMacroName = "ExcelSupport_UndoTranslation";
        public const string RedoMacroName = "ExcelSupport_RedoTranslation";

        public static void RecordAndApply(_Worksheet ws, List<CellBackupItem> items, string actionName = "Dịch Thuật AI")
        {
            if (ws == null || items == null || items.Count == 0) return;

            string? wbName = null;
            try
            {
                var wb = ws.Parent as Workbook;
                wbName = wb?.Name;
            }
            catch { }

            ExcelUndoHelper.RecordAction(new TranslationUndoAction
            {
                SheetName = ws.Name,
                WorkbookName = wbName,
                ActionName = actionName,
                Cells = items
            });
        }

        [ExcelCommand(Name = UndoMacroName)]
        public static void UndoTranslation()
        {
            ExcelUndoHelper.GlobalUndo();
        }

        [ExcelCommand(Name = RedoMacroName)]
        public static void RedoTranslation()
        {
            ExcelUndoHelper.GlobalRedo();
        }
    }
}

