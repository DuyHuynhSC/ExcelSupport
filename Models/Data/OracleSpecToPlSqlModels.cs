using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum OraclePlSqlGenMode
    {
        DdlScript,           // CREATE TABLE, COMMENT ON TABLE/COLUMN, PK, Sequence/Trigger
        PlSqlSafeBlock,      // DECLARE v_count NUMBER; BEGIN ... EXCEPTION ... END;
        PlSqlCrudProcedures, // SP_INSERT, SP_UPDATE, SP_DELETE, SP_SELECT
        PlSqlMergeUpsert,    // MERGE INTO table USING (...) ON (...)
        CSharpEntity,        // C# Entity Class with XML comments
        TypeScriptDto        // TypeScript Interface with JSDoc
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

            // Nếu type đã có dấu ngoặc tròn (ví dụ NUMBER(10,2) hay VARCHAR2(50)) thì giữ nguyên
            if (type.Contains("(") && type.Contains(")"))
            {
                return type;
            }

            return $"{type}({DataLength.Trim()})";
        }
    }

    public class OracleSpecOptions
    {
        public string SchemaName { get; set; } = string.Empty;
        public string TableName { get; set; } = "MY_TABLE";
        public string TableLogicalName { get; set; } = string.Empty;
        public OraclePlSqlGenMode GenMode { get; set; } = OraclePlSqlGenMode.DdlScript;
        public bool IncludeComments { get; set; } = true;
        public bool IncludeSequenceTrigger { get; set; } = false;
        public bool UseIdentityColumn { get; set; } = false;
        public bool IncludeSafeCheck { get; set; } = true;
        public List<OracleColumnSpec> Columns { get; set; } = new();
    }

    public class OracleSpecParseResult
    {
        public bool Success { get; set; } = true;
        public string? ErrorMessage { get; set; }
        public string DetectedTableName { get; set; } = "MY_TABLE";
        public string DetectedTableLogicalName { get; set; } = string.Empty;
        public int HeaderRowIndex { get; set; } = 1;
        public List<OracleColumnSpec> Columns { get; set; } = new();
    }
}
