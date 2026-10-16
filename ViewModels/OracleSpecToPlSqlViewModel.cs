using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.ViewModels
{
    public class OracleConditionItemViewModel : ObservableModel
    {
        private int _rowIndex;
        private string _logicalName = string.Empty;
        private string _physicalColumn = string.Empty;
        private string _operator = "=";
        private string _rawValue = string.Empty;
        private string _targetSqlValue = string.Empty;
        private string _bindVariableName = string.Empty;
        private bool _isSelected = true;

        public int RowIndex
        {
            get => _rowIndex;
            set => SetProperty(ref _rowIndex, value);
        }

        public string LogicalName
        {
            get => _logicalName;
            set => SetProperty(ref _logicalName, value);
        }

        public string PhysicalColumn
        {
            get => _physicalColumn;
            set => SetProperty(ref _physicalColumn, value);
        }

        public string Operator
        {
            get => _operator;
            set => SetProperty(ref _operator, value);
        }

        public string RawValue
        {
            get => _rawValue;
            set => SetProperty(ref _rawValue, value);
        }

        public string TargetSqlValue
        {
            get => _targetSqlValue;
            set => SetProperty(ref _targetSqlValue, value);
        }

        public string BindVariableName
        {
            get => _bindVariableName;
            set => SetProperty(ref _bindVariableName, value);
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public OracleConditionSpec ToModel()
        {
            return new OracleConditionSpec
            {
                RowIndex = RowIndex,
                LogicalName = LogicalName,
                PhysicalColumn = PhysicalColumn,
                Operator = Operator,
                RawValue = RawValue,
                TargetSqlValue = TargetSqlValue,
                BindVariableName = BindVariableName,
                IsSelected = IsSelected
            };
        }
    }

    public class OracleColumnSpecViewModel : ObservableModel
    {
        private int _rowIndex;
        private string _physicalName = string.Empty;
        private string _logicalName = string.Empty;
        private string _oracleDataType = "VARCHAR2";
        private string _dataLength = string.Empty;
        private bool _isNullable = true;
        private bool _isPrimaryKey = false;
        private string _defaultValue = string.Empty;
        private string _comments = string.Empty;
        private bool _isSelected = true;

        public int RowIndex
        {
            get => _rowIndex;
            set => SetProperty(ref _rowIndex, value);
        }

        public string PhysicalName
        {
            get => _physicalName;
            set => SetProperty(ref _physicalName, value);
        }

        public string LogicalName
        {
            get => _logicalName;
            set => SetProperty(ref _logicalName, value);
        }

        public string OracleDataType
        {
            get => _oracleDataType;
            set => SetProperty(ref _oracleDataType, value);
        }

        public string DataLength
        {
            get => _dataLength;
            set => SetProperty(ref _dataLength, value);
        }

        public bool IsNullable
        {
            get => _isNullable;
            set => SetProperty(ref _isNullable, value);
        }

        public bool IsPrimaryKey
        {
            get => _isPrimaryKey;
            set
            {
                if (SetProperty(ref _isPrimaryKey, value) && value)
                {
                    IsNullable = false;
                }
            }
        }

        public string DefaultValue
        {
            get => _defaultValue;
            set => SetProperty(ref _defaultValue, value);
        }

        public string Comments
        {
            get => _comments;
            set => SetProperty(ref _comments, value);
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public OracleColumnSpec ToModel()
        {
            return new OracleColumnSpec
            {
                RowIndex = RowIndex,
                PhysicalName = PhysicalName,
                LogicalName = LogicalName,
                OracleDataType = OracleDataType,
                DataLength = DataLength,
                IsNullable = IsNullable,
                IsPrimaryKey = IsPrimaryKey,
                DefaultValue = DefaultValue,
                Comments = Comments,
                IsSelected = IsSelected
            };
        }
    }

    public class OracleSpecToPlSqlViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private OracleSpecType _specType = OracleSpecType.QueryLogicSpec;
        private string _schemaName = string.Empty;
        private string _tableName = "MY_TABLE";
        private string _tableLogicalName = string.Empty;
        private OraclePlSqlGenMode _selectedGenMode = OraclePlSqlGenMode.QueryCheckBlock;
        private bool _includeComments = true;
        private bool _includeSequenceTrigger = false;
        private bool _useIdentityColumn = false;
        private bool _includeSafeCheck = true;

        private string _generatedScriptText = string.Empty;
        private string _statusMessage = string.Empty;
        private int _itemsCount = 0;
        private int _selectedItemsCount = 0;
        private long _executionTimeMs = 0;
        private bool _isGenerating = false;

        private OracleQuerySpecInfo _cachedQuerySpec = new();

        public OracleSpecType SpecType
        {
            get => _specType;
            set
            {
                if (SetProperty(ref _specType, value))
                {
                    OnPropertyChanged(nameof(IsQueryLogicSpec));
                    OnPropertyChanged(nameof(IsTableDefinitionDdl));
                }
            }
        }

        public bool IsQueryLogicSpec => SpecType == OracleSpecType.QueryLogicSpec;
        public bool IsTableDefinitionDdl => SpecType == OracleSpecType.TableDefinitionDdl;

        public string SchemaName
        {
            get => _schemaName;
            set
            {
                if (SetProperty(ref _schemaName, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public string TableName
        {
            get => _tableName;
            set
            {
                if (SetProperty(ref _tableName, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public string TableLogicalName
        {
            get => _tableLogicalName;
            set
            {
                if (SetProperty(ref _tableLogicalName, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public OraclePlSqlGenMode SelectedGenMode
        {
            get => _selectedGenMode;
            set
            {
                if (SetProperty(ref _selectedGenMode, value))
                {
                    OnPropertyChanged(nameof(IsModeQueryCheck));
                    OnPropertyChanged(nameof(IsModeQuerySelect));
                    OnPropertyChanged(nameof(IsModeQueryCursor));
                    OnPropertyChanged(nameof(IsModeQueryProcedure));
                    OnPropertyChanged(nameof(IsModeQueryCSharp));
                    OnPropertyChanged(nameof(IsModeDdl));
                    OnPropertyChanged(nameof(IsModePlSqlBlock));
                    OnPropertyChanged(nameof(IsModeCrudProcedures));
                    OnPropertyChanged(nameof(IsModeMergeUpsert));
                    OnPropertyChanged(nameof(IsModeCSharp));
                    OnPropertyChanged(nameof(IsModeTypeScript));
                    TriggerGenerateScript();
                }
            }
        }

        // Mode flags cho Query / Logic Spec
        public bool IsModeQueryCheck => SelectedGenMode == OraclePlSqlGenMode.QueryCheckBlock;
        public bool IsModeQuerySelect => SelectedGenMode == OraclePlSqlGenMode.QueryPureSelect;
        public bool IsModeQueryCursor => SelectedGenMode == OraclePlSqlGenMode.QueryCursorLoop;
        public bool IsModeQueryProcedure => SelectedGenMode == OraclePlSqlGenMode.QueryProcedure;
        public bool IsModeQueryCSharp => SelectedGenMode == OraclePlSqlGenMode.QueryCSharpCode;

        // Mode flags cho Table Definition DDL Spec
        public bool IsModeDdl => SelectedGenMode == OraclePlSqlGenMode.DdlScript;
        public bool IsModePlSqlBlock => SelectedGenMode == OraclePlSqlGenMode.PlSqlSafeBlock;
        public bool IsModeCrudProcedures => SelectedGenMode == OraclePlSqlGenMode.PlSqlCrudProcedures;
        public bool IsModeMergeUpsert => SelectedGenMode == OraclePlSqlGenMode.PlSqlMergeUpsert;
        public bool IsModeCSharp => SelectedGenMode == OraclePlSqlGenMode.CSharpEntity;
        public bool IsModeTypeScript => SelectedGenMode == OraclePlSqlGenMode.TypeScriptDto;

        public bool IncludeComments
        {
            get => _includeComments;
            set
            {
                if (SetProperty(ref _includeComments, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool IncludeSequenceTrigger
        {
            get => _includeSequenceTrigger;
            set
            {
                if (SetProperty(ref _includeSequenceTrigger, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool UseIdentityColumn
        {
            get => _useIdentityColumn;
            set
            {
                if (SetProperty(ref _useIdentityColumn, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public bool IncludeSafeCheck
        {
            get => _includeSafeCheck;
            set
            {
                if (SetProperty(ref _includeSafeCheck, value))
                {
                    TriggerGenerateScript();
                }
            }
        }

        public string GeneratedScriptText
        {
            get => _generatedScriptText;
            set => SetProperty(ref _generatedScriptText, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int ItemsCount
        {
            get => _itemsCount;
            set => SetProperty(ref _itemsCount, value);
        }

        public int SelectedItemsCount
        {
            get => _selectedItemsCount;
            set => SetProperty(ref _selectedItemsCount, value);
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

        public bool CanSendToOracleQuery => !string.IsNullOrWhiteSpace(GeneratedScriptText) &&
                                            (SelectedGenMode == OraclePlSqlGenMode.QueryCheckBlock ||
                                             SelectedGenMode == OraclePlSqlGenMode.QueryPureSelect ||
                                             SelectedGenMode == OraclePlSqlGenMode.QueryCursorLoop ||
                                             SelectedGenMode == OraclePlSqlGenMode.QueryProcedure ||
                                             SelectedGenMode == OraclePlSqlGenMode.DdlScript ||
                                             SelectedGenMode == OraclePlSqlGenMode.PlSqlSafeBlock ||
                                             SelectedGenMode == OraclePlSqlGenMode.PlSqlCrudProcedures ||
                                             SelectedGenMode == OraclePlSqlGenMode.PlSqlMergeUpsert);

        // Danh sách điều kiện (cho Query / Logic Spec)
        public ObservableCollection<OracleConditionItemViewModel> Conditions { get; } = new();

        // Danh sách cột (cho Table Definition DDL Spec)
        public ObservableCollection<OracleColumnSpecViewModel> Columns { get; } = new();

        public IReadOnlyList<string> CommonOracleDataTypes { get; } = new[]
        {
            "VARCHAR2",
            "NVARCHAR2",
            "CHAR",
            "NUMBER",
            "DATE",
            "TIMESTAMP",
            "CLOB",
            "BLOB"
        };

        public IReadOnlyList<string> CommonComparisonOperators { get; } = new[]
        {
            "=",
            "<>",
            "!=",
            ">=",
            "<=",
            ">",
            "<",
            "LIKE",
            "IN",
            "NOT IN",
            "IS NULL",
            "IS NOT NULL"
        };

        public ICommand GenerateScriptCommand { get; }
        public ICommand SetGenModeCommand { get; }
        public ICommand CopyClipboardCommand { get; }
        public ICommand SaveFileCommand { get; }
        public ICommand SendToOracleQuickQueryCommand { get; }
        public ICommand SelectAllItemsCommand { get; }
        public ICommand DeselectAllItemsCommand { get; }
        public ICommand RefreshFromSelectionCommand { get; }

        public Action<string>? OnSendToOracleQueryAction { get; set; }

        public OracleSpecToPlSqlViewModel(ExcelApp excelApp, bool isDarkTheme = false)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            IsDarkTheme = isDarkTheme;

            GenerateScriptCommand = new RelayCommand(_ => TriggerGenerateScript());
            SetGenModeCommand = new RelayCommand(param =>
            {
                if (param is OraclePlSqlGenMode mode)
                {
                    SelectedGenMode = mode;
                }
                else if (param is string str && Enum.TryParse<OraclePlSqlGenMode>(str, out var parsed))
                {
                    SelectedGenMode = parsed;
                }
            });
            CopyClipboardCommand = new RelayCommand(_ => ExecuteCopyClipboard());
            SaveFileCommand = new RelayCommand(_ => ExecuteSaveFile());
            SendToOracleQuickQueryCommand = new RelayCommand(_ => ExecuteSendToOracleQuery(), _ => CanSendToOracleQuery);
            SelectAllItemsCommand = new RelayCommand(_ => ExecuteSetAllItemsSelection(true));
            DeselectAllItemsCommand = new RelayCommand(_ => ExecuteSetAllItemsSelection(false));
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
                    StatusMessage = LocalizationService.Get("OracleSpec_ErrNoSelection", "⚠️ Vui lòng quét chọn một vùng đặc tả trên bảng tính Excel!");
                    return;
                }

                string sheetName = rng.Worksheet?.Name ?? "MY_TABLE";
                object[,]? values = OracleSpecToPlSqlService.GetRangeValues2D(rng);

                if (values == null)
                {
                    StatusMessage = LocalizationService.Get("OracleSpec_ErrEmptyRange", "⚠️ Vùng chọn không có dữ liệu.");
                    return;
                }

                var parseResult = OracleSpecToPlSqlService.ParseRangeSpec(values, sheetName);
                if (!parseResult.Success)
                {
                    StatusMessage = $"❌ {parseResult.ErrorMessage}";
                    return;
                }

                SpecType = parseResult.SpecType;
                _tableName = parseResult.DetectedTableName;
                OnPropertyChanged(nameof(TableName));

                _tableLogicalName = parseResult.DetectedTableLogicalName;
                OnPropertyChanged(nameof(TableLogicalName));

                if (SpecType == OracleSpecType.QueryLogicSpec)
                {
                    // Nạp Query Conditions
                    _cachedQuerySpec = parseResult.QuerySpec;
                    Conditions.Clear();
                    foreach (var cond in parseResult.QuerySpec.Conditions)
                    {
                        var vm = new OracleConditionItemViewModel
                        {
                            RowIndex = cond.RowIndex,
                            LogicalName = cond.LogicalName,
                            PhysicalColumn = cond.PhysicalColumn,
                            Operator = cond.Operator,
                            RawValue = cond.RawValue,
                            TargetSqlValue = cond.TargetSqlValue,
                            BindVariableName = cond.BindVariableName,
                            IsSelected = cond.IsSelected
                        };

                        vm.PropertyChanged += (s, e) =>
                        {
                            UpdateCounts();
                            TriggerGenerateScript();
                        };

                        Conditions.Add(vm);
                    }

                    SelectedGenMode = OraclePlSqlGenMode.QueryCheckBlock;
                    UpdateCounts();
                    TriggerGenerateScript();

                    StatusMessage = LocalizationService.Get("OracleSpec_LoadedQuerySuccess",
                        "✓ Đã nạp thành công {0} điều kiện trích xuất logic cho bảng {1}", Conditions.Count, TableName);
                }
                else
                {
                    // Nạp Table Definition DDL Columns
                    Columns.Clear();
                    foreach (var col in parseResult.Columns)
                    {
                        var vm = new OracleColumnSpecViewModel
                        {
                            RowIndex = col.RowIndex,
                            PhysicalName = col.PhysicalName,
                            LogicalName = col.LogicalName,
                            OracleDataType = col.OracleDataType,
                            DataLength = col.DataLength,
                            IsNullable = col.IsNullable,
                            IsPrimaryKey = col.IsPrimaryKey,
                            DefaultValue = col.DefaultValue,
                            Comments = col.Comments,
                            IsSelected = col.IsSelected
                        };

                        vm.PropertyChanged += (s, e) =>
                        {
                            UpdateCounts();
                            TriggerGenerateScript();
                        };

                        Columns.Add(vm);
                    }

                    SelectedGenMode = OraclePlSqlGenMode.DdlScript;
                    UpdateCounts();
                    TriggerGenerateScript();

                    StatusMessage = LocalizationService.Get("OracleSpec_LoadedSuccess",
                        "✓ Đã nạp thành công {0} cột đặc tả từ vùng chọn Excel.", Columns.Count);
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void UpdateCounts()
        {
            if (SpecType == OracleSpecType.QueryLogicSpec)
            {
                ItemsCount = Conditions.Count;
                SelectedItemsCount = Conditions.Count(c => c.IsSelected);
            }
            else
            {
                ItemsCount = Columns.Count;
                SelectedItemsCount = Columns.Count(c => c.IsSelected);
            }
        }

        public void TriggerGenerateScript()
        {
            if (IsGenerating) return;

            var sw = Stopwatch.StartNew();
            IsGenerating = true;

            try
            {
                var options = new OracleSpecOptions
                {
                    SpecType = SpecType,
                    SchemaName = SchemaName.Trim(),
                    TableName = string.IsNullOrWhiteSpace(TableName) ? "MY_TABLE" : TableName.Trim(),
                    TableLogicalName = TableLogicalName.Trim(),
                    GenMode = SelectedGenMode,
                    IncludeComments = IncludeComments,
                    IncludeSequenceTrigger = IncludeSequenceTrigger,
                    UseIdentityColumn = UseIdentityColumn,
                    IncludeSafeCheck = IncludeSafeCheck,
                    Columns = Columns.Select(c => c.ToModel()).ToList(),
                    QuerySpec = new OracleQuerySpecInfo
                    {
                        ActionType = _cachedQuerySpec.ActionType,
                        TableName = string.IsNullOrWhiteSpace(TableName) ? "MY_TABLE" : TableName.Trim(),
                        TableLogicalName = TableLogicalName.Trim(),
                        Conditions = Conditions.Select(c => c.ToModel()).ToList(),
                        OnExistsAction = _cachedQuerySpec.OnExistsAction,
                        OnNotExistsAction = _cachedQuerySpec.OnNotExistsAction,
                        OtherNotes = _cachedQuerySpec.OtherNotes
                    }
                };

                GeneratedScriptText = OracleSpecToPlSqlService.GenerateScript(options);
                sw.Stop();
                ExecutionTimeMs = sw.ElapsedMilliseconds;

                StatusMessage = LocalizationService.Get("OracleSpec_GeneratedSuccess",
                    "✓ Đã sinh thành công ({0} mục, {1}ms)", SelectedItemsCount, ExecutionTimeMs);
            }
            catch (Exception ex)
            {
                GeneratedScriptText = $"-- Error: {ex.Message}";
                StatusMessage = $"❌ {ex.Message}";
            }
            finally
            {
                IsGenerating = false;
            }
        }

        private void ExecuteCopyClipboard()
        {
            if (string.IsNullOrWhiteSpace(GeneratedScriptText)) return;

            try
            {
                System.Windows.Clipboard.SetText(GeneratedScriptText);
                StatusMessage = LocalizationService.Get("OracleSpec_CopiedToClipboard", "✓ Đã sao chép mã kịch bản vào Clipboard!");
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void ExecuteSaveFile()
        {
            if (string.IsNullOrWhiteSpace(GeneratedScriptText)) return;

            try
            {
                bool isCs = (SelectedGenMode == OraclePlSqlGenMode.CSharpEntity || SelectedGenMode == OraclePlSqlGenMode.QueryCSharpCode);
                bool isTs = (SelectedGenMode == OraclePlSqlGenMode.TypeScriptDto);

                string ext = isCs ? "cs" : (isTs ? "ts" : "sql");
                string filter = isCs ? "C# Source File (*.cs)|*.cs"
                    : isTs ? "TypeScript File (*.ts)|*.ts"
                    : "Oracle SQL Script (*.sql)|*.sql|All Files (*.*)|*.*";

                using var sfd = new SaveFileDialog
                {
                    Filter = filter,
                    FileName = $"{TableName.ToLowerInvariant()}_plsql.{ext}",
                    Title = LocalizationService.Get("OracleSpec_SaveFileTitle", "Lưu mã kịch bản Oracle")
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, GeneratedScriptText, new System.Text.UTF8Encoding(true));
                    StatusMessage = LocalizationService.Get("OracleSpec_SavedSuccess", "✓ Đã lưu file: {0}", Path.GetFileName(sfd.FileName));
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void ExecuteSendToOracleQuery()
        {
            if (string.IsNullOrWhiteSpace(GeneratedScriptText)) return;

            OnSendToOracleQueryAction?.Invoke(GeneratedScriptText);
        }

        private void ExecuteSetAllItemsSelection(bool isSelected)
        {
            if (SpecType == OracleSpecType.QueryLogicSpec)
            {
                foreach (var c in Conditions) c.IsSelected = isSelected;
            }
            else
            {
                foreach (var col in Columns) col.IsSelected = isSelected;
            }

            UpdateCounts();
            TriggerGenerateScript();
        }
    }
}
