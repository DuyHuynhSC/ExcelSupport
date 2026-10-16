using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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

        private string _schemaName = string.Empty;
        private string _tableName = "MY_TABLE";
        private string _tableLogicalName = string.Empty;
        private OraclePlSqlGenMode _selectedGenMode = OraclePlSqlGenMode.DdlScript;
        private bool _includeComments = true;
        private bool _includeSequenceTrigger = false;
        private bool _useIdentityColumn = false;
        private bool _includeSafeCheck = true;

        private string _generatedScriptText = string.Empty;
        private string _statusMessage = string.Empty;
        private int _columnsCount = 0;
        private int _selectedColumnsCount = 0;
        private long _executionTimeMs = 0;
        private bool _isGenerating = false;

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
                    OnPropertyChanged(nameof(IsDdlMode));
                    OnPropertyChanged(nameof(CanSendToOracleQuery));
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

        public int ColumnsCount
        {
            get => _columnsCount;
            set => SetProperty(ref _columnsCount, value);
        }

        public int SelectedColumnsCount
        {
            get => _selectedColumnsCount;
            set => SetProperty(ref _selectedColumnsCount, value);
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

        public bool IsDdlMode => SelectedGenMode == OraclePlSqlGenMode.DdlScript || SelectedGenMode == OraclePlSqlGenMode.PlSqlSafeBlock;
        public bool CanSendToOracleQuery => SelectedGenMode == OraclePlSqlGenMode.DdlScript ||
                                            SelectedGenMode == OraclePlSqlGenMode.PlSqlSafeBlock ||
                                            SelectedGenMode == OraclePlSqlGenMode.PlSqlCrudProcedures ||
                                            SelectedGenMode == OraclePlSqlGenMode.PlSqlMergeUpsert;

        public bool IsModeDdl => SelectedGenMode == OraclePlSqlGenMode.DdlScript;
        public bool IsModePlSqlBlock => SelectedGenMode == OraclePlSqlGenMode.PlSqlSafeBlock;
        public bool IsModeCrudProcedures => SelectedGenMode == OraclePlSqlGenMode.PlSqlCrudProcedures;
        public bool IsModeMergeUpsert => SelectedGenMode == OraclePlSqlGenMode.PlSqlMergeUpsert;
        public bool IsModeCSharp => SelectedGenMode == OraclePlSqlGenMode.CSharpEntity;
        public bool IsModeTypeScript => SelectedGenMode == OraclePlSqlGenMode.TypeScriptDto;

        public IReadOnlyList<OraclePlSqlGenMode> AvailableGenModes { get; } = new[]
        {
            OraclePlSqlGenMode.DdlScript,
            OraclePlSqlGenMode.PlSqlSafeBlock,
            OraclePlSqlGenMode.PlSqlCrudProcedures,
            OraclePlSqlGenMode.PlSqlMergeUpsert,
            OraclePlSqlGenMode.CSharpEntity,
            OraclePlSqlGenMode.TypeScriptDto
        };

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

        public ICommand GenerateScriptCommand { get; }
        public ICommand SetGenModeCommand { get; }
        public ICommand CopyClipboardCommand { get; }
        public ICommand SaveFileCommand { get; }
        public ICommand SendToOracleQuickQueryCommand { get; }
        public ICommand SelectAllColumnsCommand { get; }
        public ICommand DeselectAllColumnsCommand { get; }
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
            SendToOracleQuickQueryCommand = new RelayCommand(_ => ExecuteSendToOracleQuery(), _ => CanSendToOracleQuery && !string.IsNullOrWhiteSpace(GeneratedScriptText));
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

                _tableName = parseResult.DetectedTableName;
                OnPropertyChanged(nameof(TableName));

                _tableLogicalName = parseResult.DetectedTableLogicalName;
                OnPropertyChanged(nameof(TableLogicalName));

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

                UpdateCounts();
                TriggerGenerateScript();

                StatusMessage = LocalizationService.Get("OracleSpec_LoadedSuccess",
                    "✓ Đã nạp thành công {0} cột đặc tả từ vùng chọn Excel.", Columns.Count);
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ {ex.Message}";
            }
        }

        private void UpdateCounts()
        {
            ColumnsCount = Columns.Count;
            SelectedColumnsCount = Columns.Count(c => c.IsSelected);
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
                    SchemaName = SchemaName.Trim(),
                    TableName = string.IsNullOrWhiteSpace(TableName) ? "MY_TABLE" : TableName.Trim(),
                    TableLogicalName = TableLogicalName.Trim(),
                    GenMode = SelectedGenMode,
                    IncludeComments = IncludeComments,
                    IncludeSequenceTrigger = IncludeSequenceTrigger,
                    UseIdentityColumn = UseIdentityColumn,
                    IncludeSafeCheck = IncludeSafeCheck,
                    Columns = Columns.Select(c => c.ToModel()).ToList()
                };

                GeneratedScriptText = OracleSpecToPlSqlService.GenerateScript(options);
                sw.Stop();
                ExecutionTimeMs = sw.ElapsedMilliseconds;

                StatusMessage = LocalizationService.Get("OracleSpec_GeneratedSuccess",
                    "✓ Đã sinh thành công {0} cột ({1}ms)", SelectedColumnsCount, ExecutionTimeMs);
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
                string ext = (SelectedGenMode == OraclePlSqlGenMode.CSharpEntity) ? "cs"
                    : (SelectedGenMode == OraclePlSqlGenMode.TypeScriptDto) ? "ts" : "sql";

                string filter = (SelectedGenMode == OraclePlSqlGenMode.CSharpEntity) ? "C# Source File (*.cs)|*.cs"
                    : (SelectedGenMode == OraclePlSqlGenMode.TypeScriptDto) ? "TypeScript File (*.ts)|*.ts"
                    : "Oracle SQL Script (*.sql)|*.sql|All Files (*.*)|*.*";

                using var sfd = new SaveFileDialog
                {
                    Filter = filter,
                    FileName = $"{TableName.ToLowerInvariant()}_spec.{ext}",
                    Title = LocalizationService.Get("OracleSpec_SaveFileTitle", "Lưu mã kịch bản Oracle")
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // Lưu file kèm UTF-8 BOM chuẩn để không bị lỗi font tiếng Nhật/Việt
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

        private void ExecuteSetAllColumnsSelection(bool isSelected)
        {
            foreach (var col in Columns)
            {
                col.IsSelected = isSelected;
            }
            UpdateCounts();
            TriggerGenerateScript();
        }
    }
}
