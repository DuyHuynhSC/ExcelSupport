using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelSupport.Models
{
    public enum OracleConnectionMode
    {
        EasyConnect,
        ConnectionString
    }

    public enum OracleServiceNameType
    {
        ServiceName,
        SID
    }

    public enum OracleCompareMode
    {
        ByKeyColumns,
        Sequential
    }

    public enum OracleRowDiffStatus
    {
        Identical,
        Modified,
        MissingInA, // Only in B
        MissingInB  // Only in A
    }

    public class OracleUserCredential : INotifyPropertyChanged
    {
        private string _id = Guid.NewGuid().ToString();
        private string _username = "";
        private string _password = "";
        private string _roleOrDescription = "";
        private bool _isDefault = false;

        public string Id
        {
            get => _id;
            set { _id = value; OnPropertyChanged(nameof(Id)); }
        }

        public string Username
        {
            get => _username;
            set { _username = value; OnPropertyChanged(nameof(Username)); OnPropertyChanged(nameof(DisplayText)); }
        }

        public string Password
        {
            get => _password;
            set { _password = value; OnPropertyChanged(nameof(Password)); }
        }

        public string RoleOrDescription
        {
            get => _roleOrDescription;
            set { _roleOrDescription = value; OnPropertyChanged(nameof(RoleOrDescription)); OnPropertyChanged(nameof(DisplayText)); }
        }

        public bool IsDefault
        {
            get => _isDefault;
            set { _isDefault = value; OnPropertyChanged(nameof(IsDefault)); OnPropertyChanged(nameof(DefaultBadge)); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string DefaultBadge => IsDefault ? "⭐ (Mặc định)" : "";

        [Newtonsoft.Json.JsonIgnore]
        public string DisplayText
        {
            get
            {
                if (string.IsNullOrWhiteSpace(RoleOrDescription))
                    return Username;
                return $"{Username} ({RoleOrDescription})";
            }
        }

        public OracleUserCredential Clone()
        {
            return new OracleUserCredential
            {
                Id = Guid.NewGuid().ToString(),
                Username = this.Username,
                Password = this.Password,
                RoleOrDescription = this.RoleOrDescription,
                IsDefault = false
            };
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class OracleConnectionProfile : INotifyPropertyChanged
    {
        private string _id = Guid.NewGuid().ToString();
        private string _name = "Oracle Connection";
        private string _host = "localhost";
        private int _port = 1521;
        private string _serviceNameOrSid = "ORCL";
        private OracleServiceNameType _serviceType = OracleServiceNameType.ServiceName;
        private string _username = "";
        private string _password = "";
        private string _defaultSchema = "";
        private string _defaultTable = "";
        private string _defaultWhereClause = "";
        private bool _isDefault = false;
        private List<OracleUserCredential> _users = new List<OracleUserCredential>();
        private string? _selectedUserId;

        public string Id
        {
            get => _id;
            set { _id = value; OnPropertyChanged(nameof(Id)); }
        }

        public bool IsDefault
        {
            get => _isDefault;
            set { _isDefault = value; OnPropertyChanged(nameof(IsDefault)); OnPropertyChanged(nameof(DefaultBadge)); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string DefaultBadge => IsDefault ? "⭐ (Mặc định)" : "";

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string Host
        {
            get => _host;
            set { _host = value; OnPropertyChanged(nameof(Host)); }
        }

        public int Port
        {
            get => _port;
            set { _port = value; OnPropertyChanged(nameof(Port)); }
        }

        public string ServiceNameOrSid
        {
            get => _serviceNameOrSid;
            set { _serviceNameOrSid = value; OnPropertyChanged(nameof(ServiceNameOrSid)); }
        }

        public OracleServiceNameType ServiceType
        {
            get => _serviceType;
            set { _serviceType = value; OnPropertyChanged(nameof(ServiceType)); }
        }

        public List<OracleUserCredential> Users
        {
            get => _users;
            set
            {
                _users = value ?? new List<OracleUserCredential>();
                OnPropertyChanged(nameof(Users));
                OnPropertyChanged(nameof(DisplaySummary));
            }
        }

        public string? SelectedUserId
        {
            get => _selectedUserId;
            set { _selectedUserId = value; OnPropertyChanged(nameof(SelectedUserId)); OnPropertyChanged(nameof(DisplaySummary)); }
        }

        public string Username
        {
            get
            {
                var user = GetEffectiveUser();
                return user != null ? user.Username : _username;
            }
            set
            {
                _username = value;
                var defUser = Users.FirstOrDefault(u => u.IsDefault) ?? Users.FirstOrDefault();
                if (defUser != null)
                {
                    defUser.Username = value;
                }
                OnPropertyChanged(nameof(Username));
                OnPropertyChanged(nameof(DisplaySummary));
            }
        }

        public string Password
        {
            get
            {
                var user = GetEffectiveUser();
                return user != null ? user.Password : _password;
            }
            set
            {
                _password = value;
                var defUser = Users.FirstOrDefault(u => u.IsDefault) ?? Users.FirstOrDefault();
                if (defUser != null)
                {
                    defUser.Password = value;
                }
                OnPropertyChanged(nameof(Password));
            }
        }

        public string DefaultSchema
        {
            get => _defaultSchema;
            set { _defaultSchema = value; OnPropertyChanged(nameof(DefaultSchema)); }
        }

        public string DefaultTable
        {
            get => _defaultTable;
            set { _defaultTable = value; OnPropertyChanged(nameof(DefaultTable)); }
        }

        public string DefaultWhereClause
        {
            get => _defaultWhereClause;
            set { _defaultWhereClause = value; OnPropertyChanged(nameof(DefaultWhereClause)); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string DisplaySummary
        {
            get
            {
                string effectiveUser = GetEffectiveUser()?.Username ?? Username;
                return $"{effectiveUser}@{Host}:{Port}/{(ServiceType == OracleServiceNameType.SID ? "SID:" : "")}{ServiceNameOrSid}";
            }
        }

        public OracleUserCredential? GetEffectiveUser(string? userId = null)
        {
            if (Users == null || Users.Count == 0)
                return null;

            if (!string.IsNullOrEmpty(userId))
            {
                var match = Users.FirstOrDefault(u => u.Id == userId);
                if (match != null) return match;
            }

            if (!string.IsNullOrEmpty(SelectedUserId))
            {
                var match = Users.FirstOrDefault(u => u.Id == SelectedUserId);
                if (match != null) return match;
            }

            return Users.FirstOrDefault(u => u.IsDefault) ?? Users.FirstOrDefault();
        }

        public void EnsureDefaultUsers()
        {
            if (Users == null)
            {
                Users = new List<OracleUserCredential>();
            }

            if (Users.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(_username))
                {
                    Users.Add(new OracleUserCredential
                    {
                        Username = _username,
                        Password = _password ?? "",
                        RoleOrDescription = "Default User",
                        IsDefault = true
                    });
                }
                else
                {
                    Users.Add(new OracleUserCredential
                    {
                        Username = "XXX_USR1",
                        Password = "",
                        RoleOrDescription = "Admin / Primary User",
                        IsDefault = true
                    });
                    Users.Add(new OracleUserCredential
                    {
                        Username = "XXX_USR9",
                        Password = "",
                        RoleOrDescription = "Read-Only / Secondary User",
                        IsDefault = false
                    });
                }
            }

            if (!Users.Any(u => u.IsDefault))
            {
                Users[0].IsDefault = true;
            }
        }

        public OracleConnectionConfig ToConnectionConfig() => GetEffectiveConfig();

        public OracleConnectionConfig GetEffectiveConfig(string? userId = null)
        {
            var user = GetEffectiveUser(userId);
            return new OracleConnectionConfig
            {
                Host = this.Host,
                Port = this.Port,
                ServiceNameOrSid = this.ServiceNameOrSid,
                ServiceType = this.ServiceType,
                Username = user?.Username ?? this._username,
                Password = user?.Password ?? this._password
            };
        }

        public OracleConnectionProfile Clone()
        {
            var cloned = new OracleConnectionProfile
            {
                Id = Guid.NewGuid().ToString(),
                Name = $"{this.Name} (Copy)",
                Host = this.Host,
                Port = this.Port,
                ServiceNameOrSid = this.ServiceNameOrSid,
                ServiceType = this.ServiceType,
                _username = this._username,
                _password = this._password,
                DefaultSchema = this.DefaultSchema,
                DefaultTable = this.DefaultTable,
                DefaultWhereClause = this.DefaultWhereClause,
                IsDefault = false,
                SelectedUserId = this.SelectedUserId
            };

            foreach (var u in this.Users)
            {
                cloned.Users.Add(new OracleUserCredential
                {
                    Id = Guid.NewGuid().ToString(),
                    Username = u.Username,
                    Password = u.Password,
                    RoleOrDescription = u.RoleOrDescription,
                    IsDefault = u.IsDefault
                });
            }

            return cloned;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class OracleQuickQueryOptions
    {
        public bool IncludeTitle { get; set; } = true;
        public string TitleColorHex { get; set; } = "#2563EB";
        public bool IncludeHeaders { get; set; } = true;
        public string HeaderBgColorHex { get; set; } = "#CCFFFF";
        public bool AutoFitColumns { get; set; } = true;
        public int MaxRows { get; set; } = 0;
    }

    public class OracleQueryHistoryItem
    {
        public string Sql { get; set; } = "";
        public DateTime ExecutedAt { get; set; } = DateTime.Now;
        public int RowCount { get; set; }
        public string? ProfileName { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public string DisplayText
        {
            get
            {
                string singleLine = Regex.Replace(Sql ?? "", @"\s+", " ").Trim();
                if (singleLine.Length > 60)
                {
                    singleLine = singleLine.Substring(0, 57) + "...";
                }
                return $"{ExecutedAt:HH:mm:ss} | {singleLine}";
            }
        }
    }

    public class OracleConnectionConfig : INotifyPropertyChanged
    {
        private string _host = "localhost";
        private int _port = 1521;
        private string _serviceNameOrSid = "ORCL";
        private OracleServiceNameType _serviceType = OracleServiceNameType.ServiceName;
        private string _username = "";
        private string _password = "";
        private string _customConnectionString = "";
        private OracleConnectionMode _connectionMode = OracleConnectionMode.EasyConnect;

        public OracleConnectionMode ConnectionMode
        {
            get => _connectionMode;
            set { _connectionMode = value; OnPropertyChanged(nameof(ConnectionMode)); }
        }

        public string Host
        {
            get => _host;
            set { _host = value; OnPropertyChanged(nameof(Host)); }
        }

        public int Port
        {
            get => _port;
            set { _port = value; OnPropertyChanged(nameof(Port)); }
        }

        public string ServiceNameOrSid
        {
            get => _serviceNameOrSid;
            set { _serviceNameOrSid = value; OnPropertyChanged(nameof(ServiceNameOrSid)); }
        }

        public OracleServiceNameType ServiceType
        {
            get => _serviceType;
            set { _serviceType = value; OnPropertyChanged(nameof(ServiceType)); }
        }

        public string Username
        {
            get => _username;
            set { _username = value; OnPropertyChanged(nameof(Username)); }
        }

        public string Password
        {
            get => _password;
            set { _password = value; OnPropertyChanged(nameof(Password)); }
        }

        public string CustomConnectionString
        {
            get => _customConnectionString;
            set { _customConnectionString = value; OnPropertyChanged(nameof(CustomConnectionString)); }
        }

        public string BuildConnectionString()
        {
            if (ConnectionMode == OracleConnectionMode.ConnectionString && !string.IsNullOrWhiteSpace(CustomConnectionString))
            {
                return CustomConnectionString.Trim();
            }

            string host = string.IsNullOrWhiteSpace(Host) ? "localhost" : Host.Trim();
            int port = Port <= 0 ? 1521 : Port;
            string svc = string.IsNullOrWhiteSpace(ServiceNameOrSid) ? "ORCL" : ServiceNameOrSid.Trim();

            string dataSource;
            if (ServiceType == OracleServiceNameType.SID)
            {
                dataSource = $"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SID={svc})))";
            }
            else
            {
                dataSource = $"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={svc})))";
            }

            return $"Data Source={dataSource};User Id={Username.Trim()};Password={Password};Connection Timeout=15;Pooling=true;Min Pool Size=1;Max Pool Size=10;";
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class OracleTableColumnInfo : INotifyPropertyChanged
    {
        private bool _isSelectedKey;
        private bool _isSelectedCompare = true;

        public string ColumnName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public int DataLength { get; set; }
        public bool Nullable { get; set; }
        public bool IsPrimaryKey { get; set; }

        public bool IsSelectedKey
        {
            get => _isSelectedKey;
            set
            {
                if (_isSelectedKey != value)
                {
                    _isSelectedKey = value;
                    OnPropertyChanged(nameof(IsSelectedKey));
                }
            }
        }

        public bool IsSelectedCompare
        {
            get => _isSelectedCompare;
            set
            {
                if (_isSelectedCompare != value)
                {
                    _isSelectedCompare = value;
                    OnPropertyChanged(nameof(IsSelectedCompare));
                }
            }
        }

        public string DisplayText => IsPrimaryKey
            ? $"🔑 {ColumnName} ({DataType})"
            : $"{ColumnName} ({DataType})";

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public enum OracleReportLayout
    {
        StackedTopBottom, // Table A at top, Table B at bottom
        SideBySide        // Side by side columns
    }

    public class OracleCompareOptions
    {
        public OracleCompareMode Mode { get; set; } = OracleCompareMode.ByKeyColumns;
        public OracleReportLayout ReportLayout { get; set; } = OracleReportLayout.StackedTopBottom;
        public string HighlightColorHex { get; set; } = "#EF4444"; // Red coral default
        public List<string> SelectedKeyColumns { get; set; } = new List<string>();
        public List<string> SelectedCompareColumns { get; set; } = new List<string>();
        public string WhereClauseA { get; set; } = string.Empty;
        public string WhereClauseB { get; set; } = string.Empty;
        public string CustomQueryA { get; set; } = string.Empty;
        public string CustomQueryB { get; set; } = string.Empty;
        public bool UseCustomQuery { get; set; } = false;
        public int MaxRows { get; set; } = 0; // 0 = Unlimited
        public bool IgnoreCase { get; set; } = false;
        public bool IgnoreWhitespace { get; set; } = true;
        public bool TrimStrings { get; set; } = true;
        public bool TreatNullAsEmpty { get; set; } = true;
        public double NumericTolerance { get; set; } = 0.0;
    }

    public class OracleCellDiff
    {
        public string ColumnName { get; set; } = string.Empty;
        public object? ValueA { get; set; }
        public object? ValueB { get; set; }
        public string ValueADisplay => ValueA == null || ValueA is DBNull ? "<NULL>" : ValueA.ToString() ?? "";
        public string ValueBDisplay => ValueB == null || ValueB is DBNull ? "<NULL>" : ValueB.ToString() ?? "";
        public bool IsDifferent { get; set; }
        public string DiffSummary => IsDifferent ? $"[{ValueADisplay}] ➔ [{ValueBDisplay}]" : ValueADisplay;
    }

    public class OracleRowDiffItem
    {
        public int RowNumber { get; set; }
        public string KeyDisplay { get; set; } = string.Empty;
        public OracleRowDiffStatus Status { get; set; }
        public Dictionary<string, object?> RowValuesA { get; set; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, object?> RowValuesB { get; set; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        public List<OracleCellDiff> CellDiffs { get; set; } = new List<OracleCellDiff>();
        public List<string> DifferingColumns { get; set; } = new List<string>();

        public int DifferingColumnCount => DifferingColumns.Count;
        public string DifferingColumnsSummary => DifferingColumns.Count > 0 ? string.Join(", ", DifferingColumns) : "-";

        public string StatusBadge => Status switch
        {
            OracleRowDiffStatus.Identical => "✅ Trùng Khớp",
            OracleRowDiffStatus.Modified => $"⚠️ Sai Lệch ({DifferingColumnCount} cột)",
            OracleRowDiffStatus.MissingInA => "➕ Chỉ có ở DB B",
            OracleRowDiffStatus.MissingInB => "➖ Chỉ có ở DB A",
            _ => "-"
        };

        public string StatusColor => Status switch
        {
            OracleRowDiffStatus.Identical => "#16A34A",
            OracleRowDiffStatus.Modified => "#D97706",
            OracleRowDiffStatus.MissingInA => "#2563EB",
            OracleRowDiffStatus.MissingInB => "#DC2626",
            _ => "#64748B"
        };
    }

    public class OracleCompareResult
    {
        public string SchemaA { get; set; } = string.Empty;
        public string TableA { get; set; } = string.Empty;
        public string ConnectionNameA { get; set; } = string.Empty;
        public string SchemaB { get; set; } = string.Empty;
        public string TableB { get; set; } = string.Empty;
        public string ConnectionNameB { get; set; } = string.Empty;
        public OracleCompareOptions Options { get; set; } = new OracleCompareOptions();
        public List<string> Columns { get; set; } = new List<string>();
        public List<string> KeyColumns { get; set; } = new List<string>();
        public List<OracleRowDiffItem> DiffItems { get; set; } = new List<OracleRowDiffItem>();
        public int TotalRowsA { get; set; }
        public int TotalRowsB { get; set; }
        public int MatchCount => DiffItems.Count(r => r.Status == OracleRowDiffStatus.Identical);
        public int ModifiedCount => DiffItems.Count(r => r.Status == OracleRowDiffStatus.Modified);
        public int MissingInACount => DiffItems.Count(r => r.Status == OracleRowDiffStatus.MissingInA);
        public int MissingInBCount => DiffItems.Count(r => r.Status == OracleRowDiffStatus.MissingInB);
        public TimeSpan ExecutionTime { get; set; }
    }

    public class OracleLastCompareSession
    {
        public string? ProfileIdA { get; set; }
        public string? ProfileNameA { get; set; }
        public string? SchemaA { get; set; }
        public string? TableA { get; set; }
        public string? WhereClauseA { get; set; }

        public string? ProfileIdB { get; set; }
        public string? ProfileNameB { get; set; }
        public string? SchemaB { get; set; }
        public string? TableB { get; set; }
        public string? WhereClauseB { get; set; }

        public OracleCompareMode Mode { get; set; } = OracleCompareMode.ByKeyColumns;
        public OracleReportLayout ReportLayout { get; set; } = OracleReportLayout.StackedTopBottom;
        public string HighlightColorHex { get; set; } = "#EF4444";
        public bool IgnoreWhitespace { get; set; } = true;
        public bool IgnoreCase { get; set; } = false;
        public int MaxRows { get; set; } = 0;
        public string? TargetAddress { get; set; }
        public List<string> SelectedKeyColumns { get; set; } = new List<string>();
    }

    public class OracleTableColumnMetadata
    {
        public int ColumnId { get; set; }
        public string ColumnName { get; set; } = "";
        public string DataType { get; set; } = "";
        public int? DataLength { get; set; }
        public int? DataPrecision { get; set; }
        public int? DataScale { get; set; }
        public bool IsNullable { get; set; }
        public bool IsPrimaryKey { get; set; }
        public int PrimaryKeyPosition { get; set; }
        public string? DataDefault { get; set; }
        public string? Comments { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public string DisplayDataType
        {
            get
            {
                if (string.Equals(DataType, "NUMBER", StringComparison.OrdinalIgnoreCase) && DataPrecision.HasValue)
                {
                    return (DataScale.HasValue && DataScale.Value > 0)
                        ? $"{DataType}({DataPrecision},{DataScale})"
                        : $"{DataType}({DataPrecision})";
                }
                if ((DataType.Contains("CHAR") || DataType.Contains("RAW")) && DataLength.HasValue && DataLength.Value > 0)
                {
                    return $"{DataType}({DataLength})";
                }
                return DataType;
            }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string PkBadge => IsPrimaryKey ? "⭐ PK" : "";

        [Newtonsoft.Json.JsonIgnore]
        public string NullableText => IsNullable ? "YES" : "NO (NOT NULL)";
    }

    public class OracleTableIndexMetadata
    {
        public string IndexName { get; set; } = "";
        public string TableName { get; set; } = "";
        public string TableOwner { get; set; } = "";
        public bool IsPrimaryKey { get; set; }
        public bool IsUnique { get; set; }
        public string Status { get; set; } = "VALID";
        public string IndexType { get; set; } = "NORMAL";
        public string ColumnsDisplay { get; set; } = "";
        public List<string> ColumnNames { get; set; } = new List<string>();

        [Newtonsoft.Json.JsonIgnore]
        public string IndexCategory => IsPrimaryKey 
            ? "Index Chính (Primary Key) ⭐" 
            : (IsUnique ? "Index Phụ (Unique)" : "Index Phụ (Secondary Index)");
    }

    public class OracleTableStructureResult
    {
        public string Owner { get; set; } = "";
        public string TableName { get; set; } = "";
        public List<OracleTableColumnMetadata> Columns { get; set; } = new List<OracleTableColumnMetadata>();
        public List<OracleTableIndexMetadata> Indexes { get; set; } = new List<OracleTableIndexMetadata>();

        public string FullTableName => string.IsNullOrWhiteSpace(Owner) ? TableName : $"{Owner}.{TableName}";

        public string GenerateSelectQuery()
        {
            string targetTable = FullTableName;
            if (Columns == null || Columns.Count == 0)
            {
                return $"SELECT * FROM {targetTable}";
            }

            var colList = string.Join(",\n    ", Columns.Select(c => c.ColumnName));
            return $"SELECT \n    {colList}\nFROM {targetTable}";
        }
    }
}

