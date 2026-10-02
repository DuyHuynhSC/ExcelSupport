using System;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelSupport.Models;
using Oracle.ManagedDataAccess.Client;

namespace ExcelSupport.Services
{
    /// <summary>
    /// Deep Module trích xuất siêu dữ liệu cấu trúc Schema và thực thi SQL thuần túy từ cơ sở dữ liệu Oracle.
    /// Hoàn toàn độc lập với Excel COM Interop và Microsoft Office runtime (100% Mockable & Testable).
    /// </summary>
    public static class OracleSchemaInspector
    {
        /// <summary>
        /// Thực thi truy vấn SELECT và lấy dữ liệu DataTable thuần túy.
        /// </summary>
        public static async Task<DataTable> ExecuteQueryAsync(OracleConnectionConfig config, string sql, int maxRows = 0)
        {
            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var dt = new DataTable();
                using (var conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    string execSql = sql.Trim().TrimEnd(';');
                    if (maxRows > 0 && !execSql.ToUpperInvariant().Contains("ROWNUM"))
                    {
                        execSql = $"SELECT * FROM ({execSql}) WHERE ROWNUM <= {maxRows}";
                    }

                    using (var cmd = new OracleCommand(execSql, conn))
                    {
                        cmd.CommandTimeout = 120;
                        using (var adapter = new OracleDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
                return dt;
            });
        }

        /// <summary>
        /// Trích xuất tên bảng từ câu lệnh SQL SELECT.
        /// </summary>
        public static string ExtractTableName(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return "QUERY_RESULT";

            try
            {
                var match = Regex.Match(sql, @"\bFROM\s+([""']?(?<schema>[a-zA-Z0-9_]+)[""']?\.)?([""']?(?<table>[a-zA-Z0-9_]+)[""']?)", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string table = match.Groups["table"].Value;
                    if (!string.IsNullOrWhiteSpace(table))
                    {
                        return table.ToUpperInvariant();
                    }
                }
            }
            catch { }

            return "QUERY_RESULT";
        }

        /// <summary>
        /// Lấy toàn bộ siêu dữ liệu (Cột, Kiểu dữ liệu, Khóa chính, Chỉ mục) của một bảng hoặc view trong Oracle.
        /// </summary>
        public static async Task<OracleTableStructureResult> GetTableStructureAsync(OracleConnectionConfig config, string fullTableName)
        {
            if (string.IsNullOrWhiteSpace(fullTableName))
                throw new ArgumentException("Tên bảng không được để trống.", nameof(fullTableName));

            string connStr = config.BuildConnectionString();
            return await Task.Run(() =>
            {
                var result = new OracleTableStructureResult();
                string cleaned = fullTableName.Trim().Replace("\"", "").Replace("'", "");
                string owner = "";
                string tableName = cleaned;

                if (cleaned.Contains("."))
                {
                    var parts = cleaned.Split('.');
                    owner = parts[0].Trim().ToUpperInvariant();
                    tableName = parts[1].Trim().ToUpperInvariant();
                }
                else
                {
                    tableName = cleaned.Trim().ToUpperInvariant();
                }

                using (var conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    // Nếu chưa có owner, tự động tìm kiếm owner trong ALL_TABLES / ALL_VIEWS
                    if (string.IsNullOrEmpty(owner))
                    {
                        using (var cmdOwner = new OracleCommand(
                            @"SELECT OWNER FROM ALL_TABLES 
                              WHERE TABLE_NAME = :tbl 
                              ORDER BY CASE 
                                WHEN OWNER = SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA') THEN 0 
                                WHEN OWNER = :usr THEN 1 
                                ELSE 2 END", conn))
                        {
                            cmdOwner.Parameters.Add(new OracleParameter("tbl", tableName));
                            cmdOwner.Parameters.Add(new OracleParameter("usr", (config.Username ?? "").ToUpperInvariant()));
                            using (var rdr = cmdOwner.ExecuteReader())
                            {
                                if (rdr.Read())
                                {
                                    owner = rdr.GetString(0);
                                }
                            }
                        }

                        if (string.IsNullOrEmpty(owner))
                        {
                            using (var cmdViewOwner = new OracleCommand(
                                @"SELECT OWNER FROM ALL_VIEWS 
                                  WHERE VIEW_NAME = :tbl 
                                  ORDER BY CASE 
                                    WHEN OWNER = SYS_CONTEXT('USERENV', 'CURRENT_SCHEMA') THEN 0 
                                    WHEN OWNER = :usr THEN 1 
                                    ELSE 2 END", conn))
                            {
                                cmdViewOwner.Parameters.Add(new OracleParameter("tbl", tableName));
                                cmdViewOwner.Parameters.Add(new OracleParameter("usr", (config.Username ?? "").ToUpperInvariant()));
                                using (var rdr = cmdViewOwner.ExecuteReader())
                                {
                                    if (rdr.Read())
                                    {
                                        owner = rdr.GetString(0);
                                    }
                                }
                            }
                        }

                        if (string.IsNullOrEmpty(owner))
                        {
                            owner = (config.Username ?? "").ToUpperInvariant();
                        }
                    }

                    result.Owner = owner;
                    result.TableName = tableName;

                    // 1. Query Primary Key columns
                    var pkDict = new System.Collections.Generic.Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdPk = new OracleCommand(
                        @"SELECT cc.COLUMN_NAME, cc.POSITION
                          FROM ALL_CONSTRAINTS con
                          JOIN ALL_CONS_COLUMNS cc 
                            ON con.OWNER = cc.OWNER AND con.CONSTRAINT_NAME = cc.CONSTRAINT_NAME
                          WHERE con.CONSTRAINT_TYPE = 'P'
                            AND con.OWNER = :owner
                            AND con.TABLE_NAME = :tableName
                          ORDER BY cc.POSITION", conn))
                    {
                        cmdPk.Parameters.Add(new OracleParameter("owner", owner));
                        cmdPk.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdPk.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                string colName = rdr.GetString(0);
                                int pos = rdr.IsDBNull(1) ? 1 : Convert.ToInt32(rdr.GetValue(1));
                                pkDict[colName] = pos;
                            }
                        }
                    }

                    // 2. Query Columns & Comments
                    using (var cmdCols = new OracleCommand(
                        @"SELECT 
                            c.COLUMN_ID,
                            c.COLUMN_NAME,
                            c.DATA_TYPE,
                            c.DATA_LENGTH,
                            c.DATA_PRECISION,
                            c.DATA_SCALE,
                            c.NULLABLE,
                            c.DATA_DEFAULT,
                            cm.COMMENTS
                          FROM ALL_TAB_COLUMNS c
                          LEFT JOIN ALL_COL_COMMENTS cm 
                            ON c.OWNER = cm.OWNER AND c.TABLE_NAME = cm.TABLE_NAME AND c.COLUMN_NAME = cm.COLUMN_NAME
                          WHERE c.OWNER = :owner AND c.TABLE_NAME = :tableName
                          ORDER BY c.COLUMN_ID", conn))
                    {
                        cmdCols.Parameters.Add(new OracleParameter("owner", owner));
                        cmdCols.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdCols.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                var col = new OracleTableColumnMetadata
                                {
                                    ColumnId = rdr.IsDBNull(0) ? 0 : Convert.ToInt32(rdr.GetValue(0)),
                                    ColumnName = rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                                    DataType = rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                                    DataLength = rdr.IsDBNull(3) ? (int?)null : Convert.ToInt32(rdr.GetValue(3)),
                                    DataPrecision = rdr.IsDBNull(4) ? (int?)null : Convert.ToInt32(rdr.GetValue(4)),
                                    DataScale = rdr.IsDBNull(5) ? (int?)null : Convert.ToInt32(rdr.GetValue(5)),
                                    IsNullable = !rdr.IsDBNull(6) && rdr.GetString(6).Equals("Y", StringComparison.OrdinalIgnoreCase),
                                    DataDefault = rdr.IsDBNull(7) ? null : rdr.GetString(7)?.Trim(),
                                    Comments = rdr.IsDBNull(8) ? null : rdr.GetString(8)?.Trim()
                                };

                                if (pkDict.TryGetValue(col.ColumnName, out int pkPos))
                                {
                                    col.IsPrimaryKey = true;
                                    col.PrimaryKeyPosition = pkPos;
                                }

                                result.Columns.Add(col);
                            }
                        }
                    }

                    // 3. Query Indexes
                    var indexMap = new System.Collections.Generic.Dictionary<string, OracleTableIndexMetadata>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdIdx = new OracleCommand(
                        @"SELECT 
                            i.INDEX_NAME,
                            i.UNIQUENESS,
                            i.STATUS,
                            i.INDEX_TYPE,
                            (SELECT COUNT(*) 
                             FROM ALL_CONSTRAINTS con 
                             WHERE con.OWNER = i.OWNER 
                               AND con.CONSTRAINT_TYPE = 'P' 
                               AND con.CONSTRAINT_NAME = i.INDEX_NAME) AS IS_PK_CONSTRAINT
                          FROM ALL_INDEXES i
                          WHERE i.TABLE_OWNER = :owner AND i.TABLE_NAME = :tableName
                          ORDER BY IS_PK_CONSTRAINT DESC, i.INDEX_NAME", conn))
                    {
                        cmdIdx.Parameters.Add(new OracleParameter("owner", owner));
                        cmdIdx.Parameters.Add(new OracleParameter("tableName", tableName));
                        using (var rdr = cmdIdx.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                string idxName = rdr.GetString(0);
                                string uniqueness = rdr.IsDBNull(1) ? "NONUNIQUE" : rdr.GetString(1);
                                string status = rdr.IsDBNull(2) ? "VALID" : rdr.GetString(2);
                                string idxType = rdr.IsDBNull(3) ? "NORMAL" : rdr.GetString(3);
                                int isPkCon = rdr.IsDBNull(4) ? 0 : Convert.ToInt32(rdr.GetValue(4));

                                var idxMeta = new OracleTableIndexMetadata
                                {
                                    IndexName = idxName,
                                    TableName = tableName,
                                    TableOwner = owner,
                                    IsPrimaryKey = isPkCon > 0,
                                    IsUnique = uniqueness.Equals("UNIQUE", StringComparison.OrdinalIgnoreCase),
                                    Status = status,
                                    IndexType = idxType
                                };
                                indexMap[idxName] = idxMeta;
                                result.Indexes.Add(idxMeta);
                            }
                        }
                    }

                    // 4. Query Index Columns
                    if (indexMap.Count > 0)
                    {
                        using (var cmdIndCols = new OracleCommand(
                            @"SELECT 
                                ic.INDEX_NAME,
                                ic.COLUMN_NAME,
                                ic.COLUMN_POSITION,
                                ic.DESCEND
                              FROM ALL_IND_COLUMNS ic
                              WHERE ic.TABLE_OWNER = :owner AND ic.TABLE_NAME = :tableName
                              ORDER BY ic.INDEX_NAME, ic.COLUMN_POSITION", conn))
                        {
                            cmdIndCols.Parameters.Add(new OracleParameter("owner", owner));
                            cmdIndCols.Parameters.Add(new OracleParameter("tableName", tableName));
                            using (var rdr = cmdIndCols.ExecuteReader())
                            {
                                var groupedCols = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>(StringComparer.OrdinalIgnoreCase);
                                while (rdr.Read())
                                {
                                    string idxName = rdr.GetString(0);
                                    string colName = rdr.GetString(1);
                                    string descend = rdr.IsDBNull(3) ? "ASC" : rdr.GetString(3);
                                    string colDesc = $"{colName} {descend}".Trim();

                                    if (!groupedCols.TryGetValue(idxName, out var list))
                                    {
                                        list = new System.Collections.Generic.List<string>();
                                        groupedCols[idxName] = list;
                                    }
                                    list.Add(colDesc);

                                    if (indexMap.TryGetValue(idxName, out var idxMeta))
                                    {
                                        idxMeta.ColumnNames.Add(colName);
                                    }
                                }

                                foreach (var kvp in groupedCols)
                                {
                                    if (indexMap.TryGetValue(kvp.Key, out var idxMeta))
                                    {
                                        idxMeta.ColumnsDisplay = string.Join(", ", kvp.Value);
                                    }
                                }
                            }
                        }
                    }
                }

                return result;
            });
        }
    }
}
