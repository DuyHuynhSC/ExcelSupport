using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSupport.Models
{
    public class DuplicateGroupItem : INotifyPropertyChanged
    {
        private int _groupId;
        private int _rowIndex;
        private bool _isMaster;
        private string _keySummary = string.Empty;
        private string _rowValuesSummary = string.Empty;
        private string _sheetName = string.Empty;
        private string _workbookName = string.Empty;
        private double _similarity = 1.0;
        private string[] _rawRowValues = new string[0];

        public int GroupId
        {
            get => _groupId;
            set => SetProperty(ref _groupId, value);
        }

        public string GroupTitle => $"Nhóm {GroupId}";

        public int RowIndex
        {
            get => _rowIndex;
            set => SetProperty(ref _rowIndex, value);
        }

        public string RowDisplay => $"Dòng {RowIndex}";

        public bool IsMaster
        {
            get => _isMaster;
            set
            {
                if (SetProperty(ref _isMaster, value))
                {
                    OnPropertyChanged(nameof(RoleDescription));
                    OnPropertyChanged(nameof(RoleBadgeColor));
                }
            }
        }

        public string RoleDescription => IsMaster ? "⭐ Dòng gốc (Giữ lại)" : "🔄 Dòng trùng (Duplicate)";

        public string RoleBadgeColor => IsMaster ? "#16A34A" : "#D97706"; // Green vs Amber

        public string KeySummary
        {
            get => _keySummary;
            set => SetProperty(ref _keySummary, value);
        }

        public string RowValuesSummary
        {
            get => _rowValuesSummary;
            set => SetProperty(ref _rowValuesSummary, value);
        }

        public string SheetName
        {
            get => _sheetName;
            set => SetProperty(ref _sheetName, value);
        }

        public string WorkbookName
        {
            get => _workbookName;
            set => SetProperty(ref _workbookName, value);
        }

        public double Similarity
        {
            get => _similarity;
            set
            {
                if (SetProperty(ref _similarity, value))
                {
                    OnPropertyChanged(nameof(SimilarityPercentage));
                }
            }
        }

        public string SimilarityPercentage => $"{Similarity * 100:0.#}%";

        public string[] RawRowValues
        {
            get => _rawRowValues;
            set => SetProperty(ref _rawRowValues, value);
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
