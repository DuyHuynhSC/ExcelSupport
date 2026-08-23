using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExcelSupport.Services;

namespace ExcelSupport.Models
{
    public class ExternalSourceItem : INotifyPropertyChanged
    {
        private string _sourcePath = string.Empty;
        private string _fileName = string.Empty;
        private bool _exists;
        private int _formulaCount;
        private string _statusDisplay = string.Empty;

        public string SourcePath
        {
            get => _sourcePath;
            set { _sourcePath = value; OnPropertyChanged(); }
        }

        public string FileName
        {
            get => _fileName;
            set { _fileName = value; OnPropertyChanged(); }
        }

        public bool Exists
        {
            get => _exists;
            set { _exists = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusDisplay)); }
        }

        public int FormulaCount
        {
            get => _formulaCount;
            set { _formulaCount = value; OnPropertyChanged(); }
        }

        public string StatusDisplay
        {
            get
            {
                if (_exists)
                {
                    return LocalizationService.CurrentLanguage switch
                    {
                        AppLanguage.Japanese => "⚠️ ファイルは存在します",
                        AppLanguage.English => "⚠️ File exists on disk",
                        _ => "⚠️ File tồn tại trên máy"
                    };
                }
                else
                {
                    return LocalizationService.CurrentLanguage switch
                    {
                        AppLanguage.Japanese => "❌ ファイルが見つかりません (リンク切れ)",
                        AppLanguage.English => "❌ File not found (Broken)",
                        _ => "❌ File không tồn tại (Broken)"
                    };
                }
            }
            set { _statusDisplay = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class BrokenFormulaCellItem : INotifyPropertyChanged
    {
        private string _sheetName = string.Empty;
        private string _cellAddress = string.Empty;
        private string _formula = string.Empty;
        private string _currentValue = string.Empty;
        private string _externalSource = string.Empty;
        private bool _isBroken;
        private bool _isSelected = true;

        public string SheetName
        {
            get => _sheetName;
            set { _sheetName = value; OnPropertyChanged(); }
        }

        public string CellAddress
        {
            get => _cellAddress;
            set { _cellAddress = value; OnPropertyChanged(); }
        }

        public string FullAddress => $"{SheetName}!{CellAddress}";

        public string Formula
        {
            get => _formula;
            set { _formula = value; OnPropertyChanged(); }
        }

        public string CurrentValue
        {
            get => _currentValue;
            set { _currentValue = value; OnPropertyChanged(); }
        }

        public string ExternalSource
        {
            get => _externalSource;
            set { _externalSource = value; OnPropertyChanged(); }
        }

        public bool IsBroken
        {
            get => _isBroken;
            set { _isBroken = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public int Row { get; set; }
        public int Column { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ExternalNamedRangeItem : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _scope = string.Empty;
        private string _refersTo = string.Empty;
        private bool _isBroken;
        private bool _isSelected = true;

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        public string Scope
        {
            get => _scope;
            set { _scope = value; OnPropertyChanged(); }
        }

        public string RefersTo
        {
            get => _refersTo;
            set { _refersTo = value; OnPropertyChanged(); }
        }

        public bool IsBroken
        {
            get => _isBroken;
            set { _isBroken = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
