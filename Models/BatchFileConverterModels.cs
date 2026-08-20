using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum BatchConvertMode
    {
        ConvertFormat,      // Chuyển đổi định dạng file (.xlsx, .xls, .xlsb, .csv, .pdf)
        SplitSheetsToFiles, // Tách từng Sheet trong file thành từng file riêng
        MergeFilesToOne     // Gộp nhiều file Excel thành 1 file chứa nhiều Sheet
    }

    public enum ExcelOutputFormat
    {
        XLSX,
        XLS,
        XLSB,
        XLSM,
        CSV,
        PDF
    }

    public class BatchFileItem
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileSize { get; set; } = string.Empty;
        public string Status { get; set; } = "Chờ xử lý";
        public bool IsSuccess { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }

    public class BatchConvertOptions
    {
        public BatchConvertMode Mode { get; set; } = BatchConvertMode.ConvertFormat;
        public List<string> InputFiles { get; set; } = new List<string>();
        public string OutputDirectory { get; set; } = string.Empty;
        public ExcelOutputFormat TargetFormat { get; set; } = ExcelOutputFormat.XLSX;
        public bool OverwriteExisting { get; set; } = true;
        public string MergedFileName { get; set; } = "Gop_Cac_File_Excel.xlsx";
    }

    public class BatchConvertResult
    {
        public bool Success { get; set; }
        public int TotalFiles { get; set; }
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
