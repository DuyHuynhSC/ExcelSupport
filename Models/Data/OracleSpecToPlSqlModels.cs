using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum OracleSpecType
    {
        QueryLogicSpec,      // Đặc tả truy vấn / trích xuất logic TKCT (VD: [DATKB] = "1", 画面.処理区分[HD_HANKB])
        TableDefinitionDdl   // Đặc tả cấu trúc bảng DDL (テーブル定義書 có DataType, Length, PK...)
    }

    public enum OraclePlSqlGenMode
    {
        // --- Modes cho Query / Logic Spec (TKCT) ---
        QueryCheckBlock,     // Khối PL/SQL kiểm tra tồn tại (DECLARE v_count NUMBER; SELECT COUNT(*) INTO v_count... IF v_count > 0...)
        QueryPureSelect,     // Câu lệnh SELECT * FROM table WHERE ...
        QueryCursorLoop,     // Khối PL/SQL duyệt Cursor (DECLARE CURSOR c_... IS SELECT... FOR r IN c_... LOOP)
        QueryProcedure,      // Stored Procedure / Function kiểm tra dữ liệu
        QueryCSharpCode,     // C# Dapper / Entity Framework query string

        // --- Modes cho Table Definition Spec (DDL) ---
        DdlScript,           // CREATE TABLE, COMMENT ON TABLE/COLUMN, PK, Sequence/Trigger
        PlSqlSafeBlock,      // DECLARE v_count NUMBER; BEGIN ... EXCEPTION ... END;
        PlSqlCrudProcedures, // SP_INSERT, SP_UPDATE, SP_DELETE, SP_SELECT
        PlSqlMergeUpsert,    // MERGE INTO table USING (...) ON (...)
        CSharpEntity,        // C# Entity Class with XML comments
        TypeScriptDto        // TypeScript Interface with JSDoc
    }

    public class OracleConditionSpec
    {
        public int RowIndex { get; set; }
        public string LogicalName { get; set; } = string.Empty;        // VD: 伝票削除区分, 処理区分
        public string PhysicalColumn { get; set; } = string.Empty;     // VD: DATKB, HANKB
        public string Operator { get; set; } = "=";                    // VD: =, <>, >=, <=, LIKE, IN
        public string RawValue { get; set; } = string.Empty;           // VD: "1", 画面.処理区分[HD_HANKB]
        public string TargetSqlValue { get; set; } = string.Empty;     // VD: '1', :HD_HANKB
        public string BindVariableName { get; set; } = string.Empty;   // VD: HD_HANKB
        public bool IsSelected { get; set; } = true;
        public string Note { get; set; } = string.Empty;
    }

    public class OracleQuerySpecInfo
    {
        public string ActionType { get; set; } = "SELECT"; // SELECT, UPDATE, DELETE, INSERT
        public string TableName { get; set; } = string.Empty;
        public string TableLogicalName { get; set; } = string.Empty;
        public List<OracleConditionSpec> Conditions { get; set; } = new();
        public string OnExistsAction { get; set; } = string.Empty;
        public string OnNotExistsAction { get; set; } = string.Empty;
        public List<string> OtherNotes { get; set; } = new();
    }

    public class OracleColumnSpec
    {
        public int RowIndex { get; set; }
        public string PhysicalName { get; set; } = string.Empty;
        public string LogicalName { get; set; } = string.Empty;
        public string OracleDataType { get; set; } = "VARCHAR2";
        public string DataLength { get; set; } = string.Empty;
        public bool IsNullable { get; set; } = true;
        public bool IsPrimaryKey { get; set; } = false;
        public string DefaultValue { get; set; } = string.Empty;
        public string Comments { get; set; } = string.Empty;
        public bool IsSelected { get; set; } = true;

        public string GetFullDataTypeDefinition()
        {
            string type = OracleDataType.Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(DataLength))
            {
                return type;
            }

            if (type.Contains("(") && type.Contains(")"))
            {
                return type;
            }

            return $"{type}({DataLength.Trim()})";
        }
    }

    public class OracleSpecOptions
    {
        public OracleSpecType SpecType { get; set; } = OracleSpecType.QueryLogicSpec;
        public string SchemaName { get; set; } = string.Empty;
        public string TableName { get; set; } = "MY_TABLE";
        public string TableLogicalName { get; set; } = string.Empty;
        public OraclePlSqlGenMode GenMode { get; set; } = OraclePlSqlGenMode.QueryCheckBlock;
        public bool IncludeComments { get; set; } = true;
        public bool IncludeSequenceTrigger { get; set; } = false;
        public bool UseIdentityColumn { get; set; } = false;
        public bool IncludeSafeCheck { get; set; } = true;
        
        // DDL Columns
        public List<OracleColumnSpec> Columns { get; set; } = new();

        // Query / Logic Spec Info
        public OracleQuerySpecInfo QuerySpec { get; set; } = new();
    }

    public class OracleSpecParseResult
    {
        public bool Success { get; set; } = true;
        public string? ErrorMessage { get; set; }
        public OracleSpecType SpecType { get; set; } = OracleSpecType.QueryLogicSpec;
        
        // Table info
        public string DetectedTableName { get; set; } = "MY_TABLE";
        public string DetectedTableLogicalName { get; set; } = string.Empty;
        public int HeaderRowIndex { get; set; } = 1;

        // DDL Columns (nếu là DDL Spec)
        public List<OracleColumnSpec> Columns { get; set; } = new();

        // Query Spec Info (nếu là Query / Logic Spec)
        public OracleQuerySpecInfo QuerySpec { get; set; } = new();
    }
}
