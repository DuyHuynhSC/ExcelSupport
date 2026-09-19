using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExcelSupport.Services;

namespace ExcelSupport.Models
{
    public enum LogicalOperator
    {
        [Description("VÀ (AND)")]
        And,
        [Description("HOẶC (OR)")]
        Or
    }

    public enum FilterOperator
    {
        // Số
        [Description("> Lớn hơn")]
        GreaterThan,
        [Description(">= Lớn hơn hoặc bằng")]
        GreaterThanOrEqual,
        [Description("< Nhỏ hơn")]
        LessThan,
        [Description("<= Nhỏ hơn hoặc bằng")]
        LessThanOrEqual,
        [Description("= Bằng chính xác")]
        Equals,
        [Description("!= Khác")]
        NotEquals,
        [Description("Trong khoảng (Between)")]
        Between,
        [Description("Ngoài khoảng (Not Between)")]
        NotBetween,
        [Description("Số chẵn (Even)")]
        IsEven,
        [Description("Số lẻ (Odd)")]
        IsOdd,

        // Chuỗi
        [Description("Chứa (Contains)")]
        Contains,
        [Description("Không chứa (Does not contain)")]
        NotContains,
        [Description("Bắt đầu bằng (Starts with)")]
        StartsWith,
        [Description("Kết thúc bằng (Ends with)")]
        EndsWith,
        [Description("Khớp biểu thức chính quy (Regex)")]
        MatchesRegex,
        [Description("Chứa tiếng Việt có dấu")]
        ContainsVietnamese,
        [Description("Ô rỗng (Is Blank)")]
        IsEmpty,
        [Description("Ô không rỗng (Is Not Blank)")]
        IsNotEmpty,

        // Ngày tháng
        [Description("Hôm nay (Today)")]
        Today,
        [Description("Trong tháng này (This Month)")]
        ThisMonth,
        [Description("Trong năm nay (This Year)")]
        ThisYear
    }

    public class ColumnHeaderItem
    {
        public int ColumnIndex { get; set; } // 1-based index
        public string ColumnLetter { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;

        public string DisplayName
        {
            get
            {
                string colPrefix = LocalizationService.CurrentLanguage == AppLanguage.Japanese ? "列 "
                                 : (LocalizationService.CurrentLanguage == AppLanguage.English ? "Column " : "Cột ");
                return string.IsNullOrWhiteSpace(HeaderText)
                    ? $"{colPrefix}{ColumnLetter}"
                    : $"{ColumnLetter}: {HeaderText}";
            }
        }

        public override string ToString() => DisplayName;
    }

    public class FilterRule : INotifyPropertyChanged
    {
        private int _columnIndex = 1;
        private string _columnName = string.Empty;
        private FilterOperator _operator = FilterOperator.GreaterThan;
        private string _value1 = string.Empty;
        private string _value2 = string.Empty;
        private bool _matchCase;

        public int ColumnIndex
        {
            get => _columnIndex;
            set { _columnIndex = value; OnPropertyChanged(); }
        }

        public string ColumnName
        {
            get => _columnName;
            set { _columnName = value; OnPropertyChanged(); }
        }

        public FilterOperator Operator
        {
            get => _operator;
            set
            {
                _operator = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsBetweenOperator));
                OnPropertyChanged(nameof(NeedsValue));
            }
        }

        public string Value1
        {
            get => _value1;
            set { _value1 = value; OnPropertyChanged(); }
        }

        public string Value2
        {
            get => _value2;
            set { _value2 = value; OnPropertyChanged(); }
        }

        public bool MatchCase
        {
            get => _matchCase;
            set { _matchCase = value; OnPropertyChanged(); }
        }

        public bool IsBetweenOperator => Operator == FilterOperator.Between || Operator == FilterOperator.NotBetween;

        public bool NeedsValue => Operator != FilterOperator.IsEmpty &&
                                  Operator != FilterOperator.IsNotEmpty &&
                                  Operator != FilterOperator.IsEven &&
                                  Operator != FilterOperator.IsOdd &&
                                  Operator != FilterOperator.ContainsVietnamese &&
                                  Operator != FilterOperator.Today &&
                                  Operator != FilterOperator.ThisMonth &&
                                  Operator != FilterOperator.ThisYear;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class FilterRuleGroup : INotifyPropertyChanged
    {
        private string _groupTitle = "Nhóm Điều Kiện";
        private LogicalOperator _innerOperator = LogicalOperator.And;
        private List<FilterRule> _rules = new List<FilterRule>();

        public string GroupTitle
        {
            get => _groupTitle;
            set { _groupTitle = value; OnPropertyChanged(); }
        }

        public LogicalOperator InnerOperator
        {
            get => _innerOperator;
            set { _innerOperator = value; OnPropertyChanged(); }
        }

        public List<FilterRule> Rules
        {
            get => _rules;
            set { _rules = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class AdvancedFilterCriteria
    {
        public List<FilterRuleGroup> Groups { get; set; } = new List<FilterRuleGroup>();
        public LogicalOperator OuterOperator { get; set; } = LogicalOperator.Or;
    }

    public class BatchListFilterCriteria : INotifyPropertyChanged
    {
        private int _targetColumnIndex = 1;
        private string _targetColumnName = string.Empty;
        private string _rawPasteText = string.Empty;
        private List<string> _parsedItems = new List<string>();
        private bool _isExactMatch = true;
        private bool _matchCase = false;
        private bool _excludeList = false; // Whitelist (false) vs Blacklist (true)

        public int TargetColumnIndex
        {
            get => _targetColumnIndex;
            set { _targetColumnIndex = value; OnPropertyChanged(); }
        }

        public string TargetColumnName
        {
            get => _targetColumnName;
            set { _targetColumnName = value; OnPropertyChanged(); }
        }

        public string RawPasteText
        {
            get => _rawPasteText;
            set { _rawPasteText = value; OnPropertyChanged(); }
        }

        public List<string> ParsedItems
        {
            get => _parsedItems;
            set { _parsedItems = value; OnPropertyChanged(); }
        }

        public bool IsExactMatch
        {
            get => _isExactMatch;
            set { _isExactMatch = value; OnPropertyChanged(); }
        }

        public bool MatchCase
        {
            get => _matchCase;
            set { _matchCase = value; OnPropertyChanged(); }
        }

        public bool ExcludeList
        {
            get => _excludeList;
            set { _excludeList = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class FilterExecutionResult
    {
        public int TotalRows { get; set; }
        public int MatchedRows { get; set; }
        public int HiddenRows => TotalRows - MatchedRows;
        public double MatchPercentage => TotalRows > 0 ? (double)MatchedRows / TotalRows * 100.0 : 0;
        public string Message { get; set; } = string.Empty;
        public bool Success { get; set; }
    }
}
