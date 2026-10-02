using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.ViewModels
{
    public class SqlColumnItemViewModel : ObservableModel
    {
        private int _columnIndex;
        private string _originalHeader = string.Empty;
        private string _columnName = string.Empty;
        private SqlDataType _dataType = SqlDataType.Auto;
        private bool _isSelected = true;
        private bool _isPrimaryKey = false;
        private bool _allowNull = true;

        public int ColumnIndex
        {
            get => _columnIndex;
            set => SetProperty(ref _columnIndex, value);
        }

        public string OriginalHeader
        {
            get => _originalHeader;
            set => SetProperty(ref _originalHeader, value);
        }

        public string ColumnName
        {
            get => _columnName;
            set => SetProperty(ref _columnName, value);
        }

        public SqlDataType DataType
        {
            get => _dataType;
            set => SetProperty(ref _dataType, value);
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public bool IsPrimaryKey
        {
            get => _isPrimaryKey;
            set => SetProperty(ref _isPrimaryKey, value);
        }

        public bool AllowNull
        {
            get => _allowNull;
            set => SetProperty(ref _allowNull, value);
        }

        public SqlColumnDefinition ToModel()
        {
            return new SqlColumnDefinition
            {
                ColumnIndex = ColumnIndex,
                OriginalHeader = OriginalHeader,
                ColumnName = ColumnName,
                DataType = DataType,
                IsSelected = IsSelected,
                IsPrimaryKey = IsPrimaryKey,
                AllowNull = AllowNull
            };
        }
    }

    public class SqlScriptGeneratorViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;
        private object[,]? _cachedValues;

        private string _tableName = "MY_TABLE";
        private string _schemaName = string.Empty;
        private SqlDialect _selectedDialect = SqlDialect.Oracle;
        private SqlStatementType _selectedStatementType = SqlStatementType.InsertSingle;
        private bool _firstRowIsHeader = true;
        private int _batchSize = 500;
        private bool _treatBlankAsNull = true;
        private bool _includeTransaction = false;
        private bool _includeIdentityInsert = false;
        private bool _useUnicodeNPrefix = true;

        private string _sqlScriptText = string.Empty;
        private string _statusMessage = string.Empty;
        private int _rowCount = 0;
        private int _columnCount = 0;
        private int _statementCount = 0;
        private long _executionTimeMs = 0;
        private bool _isGenerating = false;

        public string TableName
        {
            get => _tableName;
            set => SetProperty(ref _tableName, value);
        }

        public string SchemaName
        {
            get => _schemaName;
            set => SetProperty(ref _schemaName, value);
        }

        public SqlDialect SelectedDialect
        {
            get => _selectedDialect;
            set
            {
                if (SetProperty(ref _selectedDialect, value))
                {
                    OnPropertyChanged(nameof(CanSendToOracleQuery));
                    OnPropertyChanged(nameof(IsIdentityInsertVisible));
                    TriggerGenerateScript();
                }
            }
        }

        public SqlStatementType SelectedStatementType
        {
            get => _selectedStatementType;
            set
            {
                if (SetProperty(ref _selectedStatementType, value))
                {
                    OnPropertyChanged(nameof(IsBatchSizeVisible));
                    OnPropertyChanged(nameof(IsMergeOrUpdate));
                    TriggerGenerateScript();
                }
            }
        }

        public bool FirstRowIsHeader
        {
            get => _firstRowIsHeader;
            set
            {
                if (SetProperty(ref _firstRowIsHeader, value))
                {
                    ReloadColumnsFromCache();
                    TriggerGenerateScript();
                }
            }
        }

        public int BatchSize
        {
            get => _batchSize;
            set
            {
                if (SetProperty(ref _batchSize, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool TreatBlankAsNull
        {
            get => _treatBlankAsNull;
            set
            {
                if (SetProperty(ref _treatBlankAsNull, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool IncludeTransaction
        {
            get => _includeTransaction;
            set
            {
                if (SetProperty(ref _includeTransaction, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool IncludeIdentityInsert
        {
            get => _includeIdentityInsert;
            set
            {
                if (SetProperty(ref _includeIdentityInsert, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool UseUnicodeNPrefix
        {
            get => _useUnicodeNPrefix;
            set
            {
                if (SetProperty(ref _useUnicodeNPrefix, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public string SqlScriptText
        {
            get => _sqlScriptText;
            set => SetProperty(ref _sqlScriptText, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int RowCount
        {
            get => _rowCount;
            set => SetProperty(ref _rowCount, value);
        }

        public int ColumnCount
        {
            get => _columnCount;
            set => SetProperty(ref _columnCount, value);
        }

        public int StatementCount
        {
            get => _statementCount;
            set => SetProperty(ref _statementCount, value);
        }

        public long ExecutionTimeMs
        {
            get => _executionTimeMs;
            set => SetProperty(ref _executionTimeMs, value);
        }

        public bool IsGenerating
        {
            get => _isGenerating;
            set => SetProperty(ref _isGenerating, value);
        }

        public bool CanSendToOracleQuery => SelectedDialect == SqlDialect.Oracle;
        public bool IsIdentityInsertVisible => SelectedDialect == SqlDialect.SqlServer;
        public bool IsBatchSizeVisible => SelectedStatementType == SqlStatementType.InsertBatch;
        public bool IsMergeOrUpdate => SelectedStatementType == SqlStatementType.MergeUpsert || SelectedStatementType == SqlStatementType.Update;

        public ObservableCollection<SqlColumnItemViewModel> Columns { get; } = new();

        public IReadOnlyList<SqlDialect> AvailableDialects { get; } = new[]
        {
            SqlDialect.Oracle,
            SqlDialect.SqlServer,
            SqlDialect.PostgreSql,
            SqlDialect.MySql,
            SqlDialect.Sqlite,
            SqlDialect.Generic
        };

        public IReadOnlyList<SqlStatementType> AvailableStatementTypes { get; } = new[]
        {
            SqlStatementType.InsertSingle,
            SqlStatementType.InsertBatch,
            SqlStatementType.MergeUpsert,
            SqlStatementType.Update
        };

        public IReadOnlyList<SqlDataType> AvailableDataTypes { get; } = new[]
        {
            SqlDataType.Auto,
            SqlDataType.Text,
            SqlDataType.Number,
            SqlDataType.Date,
            SqlDataType.Timestamp,
            SqlDataType.Boolean,
            SqlDataType.Raw
        };

        public ICommand GenerateScriptCommand { get; }
        public ICommand CopyClipboardCommand { get; }
        public ICommand SaveFileCommand { get; }
        public ICommand SendToOracleQuickQueryCommand { get; }
        public ICommand SelectAllColumnsCommand { get; }
        public ICommand DeselectAllColumnsCommand { get; }
        public ICommand RefreshFromSelectionCommand { get; }

        public Action<string>? OnSendToOracleQueryAction { get; set; }

        public SqlScriptGeneratorViewModel(ExcelApp excelApp, bool isDarkTheme = false)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            IsDarkTheme = isDarkTheme;

            GenerateScriptCommand = new RelayCommand(_ => TriggerGenerateScript());
            CopyClipboardCommand = new RelayCommand(_ => ExecuteCopyClipboard());
            SaveFileCommand = new RelayCommand(_ => ExecuteSaveFile());
            SendToOracleQuickQueryCommand = new RelayCommand(_ => ExecuteSendToOracleQuery(), _ => CanSendToOracleQuery && !string.IsNullOrWhiteSpace(SqlScriptText));
            SelectAllColumnsCommand = new RelayCommand(_ => ExecuteSetAllColumnsSelection(true));
            DeselectAllColumnsCommand = new RelayCommand(_ => ExecuteSetAllColumnsSelection(false));
            RefreshFromSelectionCommand = new RelayCommand(_ => LoadFromCurrentSelection());

            LoadFromCurrentSelection();
        }

        public void LoadFromCurrentSelection()
        {
            try
            {
                dynamic sel = _excelApp.Selection;
                if (sel is not Range rng)
                {
                    StatusMessage = LocalizationService.Get("SqlGen_ErrNoSelection", "⚠️ Vui lòng chọn một vùng ô trên bảng tính Excel!");
                    return;
                }

                // Gợi ý tên bảng theo tên Sheet hiện tại
                try
                {
                    string sheetName = rng.Worksheet?.Name ?? "MY_TABLE";
                    TableName = SqlScriptGeneratorService.SanitizeColumnName(sheetName);
                }
                catch
                {
                    TableName = "MY_TABLE";
                }

                _cachedValues = SqlScriptGeneratorService.GetRangeValues2D(rng);
                ReloadColumnsFromCache();
                TriggerGenerateScript();
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void ReloadColumnsFromCache()
        {
            if (_cachedValues == null) return;

            int rows = _cachedValues.GetLength(0);
            int cols = _cachedValues.GetLength(1);

            Columns.Clear();
            int dataStartRow = FirstRowIsHeader ? 2 : 1;

            for (int c = 1; c <= cols; c++)
            {
                string headerName = string.Empty;
                if (FirstRowIsHeader)
                {
                    object? hVal = _cachedValues[1, c];
                    headerName = hVal?.ToString()?.Trim() ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(headerName))
                {
                    headerName = $"COL{c}";
                }

                string sanitized = SqlScriptGeneratorService.SanitizeColumnName(headerName);

                bool isPkCandidate = c == 1 && (
                    sanitized.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) ||
                    sanitized.Equals("ID", StringComparison.OrdinalIgnoreCase) ||
                    sanitized.EndsWith("_CODE", StringComparison.OrdinalIgnoreCase) ||
                    sanitized.Equals("CODE", StringComparison.OrdinalIgnoreCase)
                );

                var item = new SqlColumnItemViewModel
                {
                    ColumnIndex = c,
                    OriginalHeader = headerName,
                    ColumnName = sanitized,
                    DataType = SqlDataType.Auto,
                    IsSelected = true,
                    IsPrimaryKey = isPkCandidate,
                    AllowNull = true
                };

                item.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(SqlColumnItemViewModel.IsSelected) ||
                        e.PropertyName == nameof(SqlColumnItemViewModel.IsPrimaryKey) ||
                        e.PropertyName == nameof(SqlColumnItemViewModel.DataType) ||
                        e.PropertyName == nameof(SqlColumnItemViewModel.ColumnName))
                    {
                        TriggerGenerateScript();
                    }
                };

                Columns.Add(item);
            }
        }

        public void TriggerGenerateScript()
        {
            if (_cachedValues == null) return;

            var options = new SqlScriptGeneratorOptions
            {
                TableName = TableName,
                SchemaName = SchemaName,
                Dialect = SelectedDialect,
                StatementType = SelectedStatementType,
                FirstRowIsHeader = FirstRowIsHeader,
                BatchSize = BatchSize,
                TreatBlankAsNull = TreatBlankAsNull,
                IncludeTransaction = IncludeTransaction,
                IncludeIdentityInsert = IncludeIdentityInsert,
                UseUnicodeNPrefix = UseUnicodeNPrefix,
                Columns = Columns.Select(c => c.ToModel()).ToList()
            };

            IsGenerating = true;
            StatusMessage = LocalizationService.Get("SqlGen_StatusGenerating", "Đang sinh script SQL...");

            Task.Run(() =>
            {
                var result = SqlScriptGeneratorService.GenerateScriptFromValues(_cachedValues, options);
                return result;
            }).ContinueWith(t =>
            {
                IsGenerating = false;
                if (t.IsFaulted)
                {
                    StatusMessage = $"❌ {t.Exception?.GetBaseException().Message}";
                    return;
                }

                var res = t.Result;
                if (!res.Success)
                {
                    SqlScriptText = string.Empty;
                    StatusMessage = $"⚠️ {res.ErrorMessage}";
                    RowCount = 0;
                    ColumnCount = 0;
                    StatementCount = 0;
                    ExecutionTimeMs = 0;
                }
                else
                {
                    SqlScriptText = res.SqlScript;
                    RowCount = res.RowCount;
                    ColumnCount = res.ColumnCount;
                    StatementCount = res.StatementCount;
                    ExecutionTimeMs = res.ExecutionTimeMs;
                    StatusMessage = LocalizationService.Get("SqlGen_StatusDone", "Đã sinh {0} câu lệnh SQL ({1} dòng, {2} ms)!", res.StatementCount, res.RowCount, res.ExecutionTimeMs);
                }
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void ExecuteCopyClipboard()
        {
            if (string.IsNullOrEmpty(SqlScriptText))
            {
                StatusMessage = LocalizationService.Get("SqlGen_ErrEmptyScript", "Không có nội dung script để sao chép!");
                return;
            }

            try
            {
                CopyToClipboard(SqlScriptText);
                StatusMessage = LocalizationService.Get("SqlGen_CopiedToClipboard", "📋 Đã sao chép toàn bộ script SQL vào Clipboard!");
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void ExecuteSaveFile()
        {
            if (string.IsNullOrEmpty(SqlScriptText))
            {
                StatusMessage = LocalizationService.Get("SqlGen_ErrEmptyScript", "Không có nội dung script để lưu file!");
                return;
            }

            try
            {
                using var sfd = new SaveFileDialog
                {
                    Filter = "SQL Files (*.sql)|*.sql|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                    DefaultExt = "sql",
                    FileName = $"{TableName}_{SelectedStatementType}_{DateTime.Now:yyyyMMdd}.sql",
                    Title = LocalizationService.Get("SqlGen_SaveFileDialogTitle", "Lưu Script SQL")
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, SqlScriptText, System.Text.Encoding.UTF8);
                    StatusMessage = LocalizationService.Get("SqlGen_FileSavedSuccess", "💾 Đã lưu file thành công tại: {0}", Path.GetFileName(sfd.FileName));
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void ExecuteSendToOracleQuery()
        {
            if (string.IsNullOrWhiteSpace(SqlScriptText)) return;
            OnSendToOracleQueryAction?.Invoke(SqlScriptText);
        }

        private void ExecuteSetAllColumnsSelection(bool isSelected)
        {
            foreach (var col in Columns)
            {
                col.IsSelected = isSelected;
            }
            TriggerGenerateScript();
        }
    }
}
