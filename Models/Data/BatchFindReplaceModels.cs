using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelSupport.Models
{
    public enum FindReplaceScope
    {
        Selection,
        ActiveSheet,
        AllSheetsCurrentWorkbook,
        AllOpenWorkbooks
    }

    public enum FindReplaceLookIn
    {
        Values,
        Formulas
    }

    public class FindReplacePair
    {
        public string FindText { get; set; } = string.Empty;
        public string ReplaceText { get; set; } = string.Empty;
        public int MatchCount { get; set; } = 0;
    }

    public class BatchFindReplaceOptions
    {
        public FindReplaceScope Scope { get; set; } = FindReplaceScope.ActiveSheet;
        public bool MatchEntireCell { get; set; } = false;
        public bool MatchCase { get; set; } = false;
        public FindReplaceLookIn LookIn { get; set; } = FindReplaceLookIn.Values;
        public bool HighlightReplacedCells { get; set; } = true;
        public Color HighlightColor { get; set; } = Color.FromArgb(254, 240, 138); // Vàng highlight dịu
        public List<FindReplacePair> Pairs { get; set; } = new List<FindReplacePair>();
    }

    public class BatchFindReplaceResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int TotalReplacements { get; set; }
        public int TotalCellsModified { get; set; }
        public int SheetsModified { get; set; }
        public List<FindReplacePair> PairResults { get; set; } = new List<FindReplacePair>();
    }
}
