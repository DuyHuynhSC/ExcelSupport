using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSupport.Models
{
    public enum DuplicateMatchMode
    {
        ExactMatch, // Trùng khớp chính xác 100%
        FuzzyMatch  // So khớp mờ (tương đồng theo tỷ lệ %)
    }

    public class ColumnSelectionItem : INotifyPropertyChanged
    {
        private int _columnIndex;
        private string _columnLetter = string.Empty;
        private string _headerName = string.Empty;
        private bool _isSelected = true;

        public int ColumnIndex
        {
            get => _columnIndex;
            set => SetProperty(ref _columnIndex, value);
        }

        public string ColumnLetter
        {
            get => _columnLetter;
            set => SetProperty(ref _columnLetter, value);
        }

        public string HeaderName
        {
            get => _headerName;
            set => SetProperty(ref _headerName, value);
        }

        public string DisplayName => string.IsNullOrEmpty(HeaderName) 
            ? $"Cột {ColumnLetter}" 
            : $"Cột {ColumnLetter}: {HeaderName}";

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }
    }

    public class DuplicateFinderOptions
    {
        public DuplicateMatchMode Mode { get; set; } = DuplicateMatchMode.ExactMatch;
        public double FuzzySimilarityThreshold { get; set; } = 0.85; // 85%
        public bool FirstRowIsHeader { get; set; } = true;
        public bool IgnoreWhitespace { get; set; } = true;
        public bool CaseInsensitive { get; set; } = true;
        public List<int> SelectedColumnIndices { get; set; } = new List<int>();
    }
}
