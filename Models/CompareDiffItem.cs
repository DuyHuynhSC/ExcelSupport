using System;
using System.Collections.Generic;
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

    public class TextDiffSegment
    {
        public string Text { get; set; } = string.Empty;
        public bool IsDiff { get; set; } = false;

        public TextDiffSegment() { }
        public TextDiffSegment(string text, bool isDiff = false)
        {
            Text = text;
            IsDiff = isDiff;
        }
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
        private List<TextDiffSegment> _oldValueSegments = new List<TextDiffSegment>();
        private List<TextDiffSegment> _newValueSegments = new List<TextDiffSegment>();

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
                    ComputeDiffSegments();
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
            set
            {
                if (SetProperty(ref _oldValue, value))
                {
                    ComputeDiffSegments();
                }
            }
        }

        public string NewValue
        {
            get => _newValue;
            set
            {
                if (SetProperty(ref _newValue, value))
                {
                    ComputeDiffSegments();
                }
            }
        }

        public List<TextDiffSegment> OldValueSegments
        {
            get => _oldValueSegments;
            private set => SetProperty(ref _oldValueSegments, value);
        }

        public List<TextDiffSegment> NewValueSegments
        {
            get => _newValueSegments;
            private set => SetProperty(ref _newValueSegments, value);
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

        public void ComputeDiffSegments()
        {
            if (Type == DiffType.Added)
            {
                OldValueSegments = new List<TextDiffSegment> { new TextDiffSegment(_oldValue, false) };
                NewValueSegments = new List<TextDiffSegment> { new TextDiffSegment(_newValue, true) };
                return;
            }
            if (Type == DiffType.Deleted)
            {
                OldValueSegments = new List<TextDiffSegment> { new TextDiffSegment(_oldValue, true) };
                NewValueSegments = new List<TextDiffSegment> { new TextDiffSegment(_newValue, false) };
                return;
            }

            string s1 = _oldValue ?? string.Empty;
            string s2 = _newValue ?? string.Empty;

            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
            {
                OldValueSegments = new List<TextDiffSegment> { new TextDiffSegment(s1, !string.IsNullOrEmpty(s1)) };
                NewValueSegments = new List<TextDiffSegment> { new TextDiffSegment(s2, !string.IsNullOrEmpty(s2)) };
                return;
            }

            int m = s1.Length;
            int n = s2.Length;

            if (m > 800 || n > 800)
            {
                int prefix = 0;
                while (prefix < m && prefix < n && s1[prefix] == s2[prefix]) prefix++;

                int suffix = 0;
                while (suffix < (m - prefix) && suffix < (n - prefix) && s1[m - 1 - suffix] == s2[n - 1 - suffix]) suffix++;

                var segs1 = new List<TextDiffSegment>();
                if (prefix > 0) segs1.Add(new TextDiffSegment(s1.Substring(0, prefix), false));
                if (m - prefix - suffix > 0) segs1.Add(new TextDiffSegment(s1.Substring(prefix, m - prefix - suffix), true));
                if (suffix > 0) segs1.Add(new TextDiffSegment(s1.Substring(m - suffix), false));

                var segs2 = new List<TextDiffSegment>();
                if (prefix > 0) segs2.Add(new TextDiffSegment(s2.Substring(0, prefix), false));
                if (n - prefix - suffix > 0) segs2.Add(new TextDiffSegment(s2.Substring(prefix, n - prefix - suffix), true));
                if (suffix > 0) segs2.Add(new TextDiffSegment(s2.Substring(n - suffix), false));

                OldValueSegments = segs1;
                NewValueSegments = segs2;
                return;
            }

            int[,] dp = new int[m + 1, n + 1];
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (s1[i] == s2[j]) dp[i + 1, j + 1] = dp[i, j] + 1;
                    else dp[i + 1, j + 1] = Math.Max(dp[i + 1, j], dp[i, j + 1]);
                }
            }

            bool[] inLcs1 = new bool[m];
            bool[] inLcs2 = new bool[n];

            int ci = m;
            int cj = n;
            while (ci > 0 && cj > 0)
            {
                if (s1[ci - 1] == s2[cj - 1])
                {
                    inLcs1[ci - 1] = true;
                    inLcs2[cj - 1] = true;
                    ci--;
                    cj--;
                }
                else if (dp[ci, cj - 1] >= dp[ci - 1, cj])
                {
                    cj--;
                }
                else
                {
                    ci--;
                }
            }

            var result1 = new List<TextDiffSegment>();
            int start1 = 0;
            while (start1 < m)
            {
                bool isMatch = inLcs1[start1];
                int end1 = start1 + 1;
                while (end1 < m && inLcs1[end1] == isMatch) end1++;
                result1.Add(new TextDiffSegment(s1.Substring(start1, end1 - start1), !isMatch));
                start1 = end1;
            }

            var result2 = new List<TextDiffSegment>();
            int start2 = 0;
            while (start2 < n)
            {
                bool isMatch = inLcs2[start2];
                int end2 = start2 + 1;
                while (end2 < n && inLcs2[end2] == isMatch) end2++;
                result2.Add(new TextDiffSegment(s2.Substring(start2, end2 - start2), !isMatch));
                start2 = end2;
            }

            OldValueSegments = result1.Count > 0 ? result1 : new List<TextDiffSegment> { new TextDiffSegment(s1, false) };
            NewValueSegments = result2.Count > 0 ? result2 : new List<TextDiffSegment> { new TextDiffSegment(s2, false) };
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
