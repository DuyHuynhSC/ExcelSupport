using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSupport.Models
{
    public enum DiffType
    {
        Modified, // Ô bị thay đổi giá trị
        Added,    // Ô hoặc Dòng chỉ có ở File B (Thêm mới)
        Deleted   // Ô hoặc Dòng chỉ có ở File A (Đã bị xóa ở File B)
    }

    public class CompareDiffItem : INotifyPropertyChanged
    {
        private int _index;
        private string _sheetName = string.Empty;
        private string _cellAddress = string.Empty;
        private string _keyIdentifier = string.Empty;
        private DiffType _type = DiffType.Modified;
        private string _oldValue = string.Empty;
        private string _newValue = string.Empty;
        private string _workbook1Name = string.Empty;
        private string _workbook2Name = string.Empty;

        public int Index
        {
            get => _index;
            set => SetProperty(ref _index, value);
        }

        public string SheetName
        {
            get => _sheetName;
            set => SetProperty(ref _sheetName, value);
        }

        public string CellAddress
        {
            get => _cellAddress;
            set => SetProperty(ref _cellAddress, value);
        }

        public string KeyIdentifier
        {
            get => _keyIdentifier;
            set => SetProperty(ref _keyIdentifier, value);
        }

        public DiffType Type
        {
            get => _type;
            set
            {
                if (SetProperty(ref _type, value))
                {
                    OnPropertyChanged(nameof(TypeDescription));
                    OnPropertyChanged(nameof(TypeBadgeColor));
                }
            }
        }

        public string TypeDescription => Type switch
        {
            DiffType.Modified => "Thay đổi",
            DiffType.Added => "Thêm mới (File B)",
            DiffType.Deleted => "Đã xóa (File A)",
            _ => "Khác"
        };

        public string TypeBadgeColor => Type switch
        {
            DiffType.Modified => "#D97706", // Amber / Cam
            DiffType.Added => "#16A34A",    // Green / Xanh lá
            DiffType.Deleted => "#DC2626",  // Red / Đỏ
            _ => "#64748B"
        };

        public string OldValue
        {
            get => _oldValue;
            set => SetProperty(ref _oldValue, value);
        }

        public string NewValue
        {
            get => _newValue;
            set => SetProperty(ref _newValue, value);
        }

        public string Workbook1Name
        {
            get => _workbook1Name;
            set => SetProperty(ref _workbook1Name, value);
        }

        public string Workbook2Name
        {
            get => _workbook2Name;
            set => SetProperty(ref _workbook2Name, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
