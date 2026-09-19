using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum FormulaErrorType
    {
        NA,             // #N/A
        Value,          // #VALUE!
        Ref,            // #REF!
        Div0,           // #DIV/0!
        Name,           // #NAME?
        Num,            // #NUM!
        Null,           // #NULL!
        Calc,           // #CALC!
        Spill,          // #SPILL!
        Unknown
    }

    public class FormulaCellItem
    {
        public string CellAddress { get; set; } = string.Empty; // e.g. "B2" or "'Sheet1'!$B$2"
        public int Row { get; set; }
        public int Column { get; set; }
        public string SheetName { get; set; } = string.Empty;
        public string Formula { get; set; } = string.Empty;
        public string DisplayValue { get; set; } = string.Empty;
        public FormulaErrorType ErrorType { get; set; } = FormulaErrorType.Unknown;
        public string ErrorTypeName => GetErrorDisplay(ErrorType);
        
        // Context around the cell
        public string HeaderText { get; set; } = string.Empty;
        public List<string> PrecedentValues { get; set; } = new List<string>();
        
        // Diagnosis & Proposed Fix
        public string? AiDiagnosis { get; set; }
        public string? ProposedFormula { get; set; }
        public string? FixExplanation { get; set; }
        public bool IsFixed { get; set; }

        public static string GetErrorDisplay(FormulaErrorType err)
        {
            return err switch
            {
                FormulaErrorType.NA => "#N/A",
                FormulaErrorType.Value => "#VALUE!",
                FormulaErrorType.Ref => "#REF!",
                FormulaErrorType.Div0 => "#DIV/0!",
                FormulaErrorType.Name => "#NAME?",
                FormulaErrorType.Num => "#NUM!",
                FormulaErrorType.Null => "#NULL!",
                FormulaErrorType.Calc => "#CALC!",
                FormulaErrorType.Spill => "#SPILL!",
                _ => "#ERROR!"
            };
        }
    }

    public class FormulaDoctorScanResult
    {
        public string WorkbookName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public int TotalCellsScanned { get; set; }
        public int TotalErrorsFound { get; set; }
        public List<FormulaCellItem> ErrorItems { get; set; } = new List<FormulaCellItem>();
        public TimeSpan ScanDuration { get; set; }
    }

    public class FormulaExplainResult
    {
        public string OriginalFormula { get; set; } = string.Empty;
        public string OverallPurpose { get; set; } = string.Empty;
        public List<FormulaStepExplanation> Steps { get; set; } = new List<FormulaStepExplanation>();
        public string ReturnTypeInfo { get; set; } = string.Empty;
        public string OptimizationAdvice { get; set; } = string.Empty;
    }

    public class FormulaStepExplanation
    {
        public int StepNumber { get; set; }
        public string SubExpression { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class FormulaModernizeResult
    {
        public string OriginalFormula { get; set; } = string.Empty;
        public string ModernizedFormula { get; set; } = string.Empty;
        public string ChangesSummary { get; set; } = string.Empty;
        public bool IsModernized { get; set; }
    }
}
