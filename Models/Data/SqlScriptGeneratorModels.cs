using System;
using System.Collections.Generic;

namespace ExcelSupport.Models
{
    public enum SqlDialect
    {
        Oracle,
        SqlServer,
        PostgreSql,
        MySql,
        Sqlite,
        Generic
    }

    public enum SqlStatementType
    {
        InsertSingle,   // INSERT INTO table (...) VALUES (...); từng dòng
        InsertBatch,    // Multi-row INSERT VALUES (...), (...); hoặc INSERT ALL (Oracle)
        MergeUpsert,    // MERGE INTO (Oracle, SQL Server) hoặc ON CONFLICT / ON DUPLICATE KEY (PG, MySQL, SQLite)
        Update          // UPDATE table SET ... WHERE pk = ...;
    }

    public enum SqlDataType
    {
        Auto,
        Text,
        Number,
        Date,
        Timestamp,
        Boolean,
        Raw
    }

    public class SqlColumnDefinition
    {
        public int ColumnIndex { get; set; }
        public string OriginalHeader { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public SqlDataType DataType { get; set; } = SqlDataType.Auto;
        public bool IsSelected { get; set; } = true;
        public bool IsPrimaryKey { get; set; } = false;
        public bool AllowNull { get; set; } = true;
    }

    public class SqlScriptGeneratorOptions
    {
        public string TableName { get; set; } = "MY_TABLE";
        public string SchemaName { get; set; } = string.Empty;
        public SqlDialect Dialect { get; set; } = SqlDialect.Oracle;
        public SqlStatementType StatementType { get; set; } = SqlStatementType.InsertSingle;
        public bool FirstRowIsHeader { get; set; } = true;
        public int BatchSize { get; set; } = 500;
        public bool TreatBlankAsNull { get; set; } = true;
        public bool IncludeTransaction { get; set; } = false;
        public bool IncludeIdentityInsert { get; set; } = false;
        public bool UseUnicodeNPrefix { get; set; } = true;
        public List<SqlColumnDefinition> Columns { get; set; } = new();
    }

    public class SqlScriptGeneratorResult
    {
        public bool Success { get; set; } = true;
        public string? ErrorMessage { get; set; }
        public string SqlScript { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public int StatementCount { get; set; }
        public long ExecutionTimeMs { get; set; }
    }
}
