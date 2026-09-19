using System;
using System.Collections.ObjectModel;
using ExcelSupport.ViewModels;

namespace ExcelSupport.Models
{
    public class ChatSheetMessageItem : ViewModelBase
    {
        private string _content = string.Empty;
        private string? _suggestedFormula;
        private bool _isUser;

        public string Id { get; set; } = Guid.NewGuid().ToString();

        public bool IsUser
        {
            get => _isUser;
            set => SetProperty(ref _isUser, value);
        }

        public string Content
        {
            get => _content;
            set => SetProperty(ref _content, value);
        }

        public DateTime Timestamp { get; set; } = DateTime.Now;

        public ObservableCollection<string> ReferencedCells { get; } = new ObservableCollection<string>();

        public bool HasReferencedCells => ReferencedCells.Count > 0;

        public string? SuggestedFormula
        {
            get => _suggestedFormula;
            set
            {
                if (SetProperty(ref _suggestedFormula, value))
                {
                    OnPropertyChanged(nameof(HasSuggestedFormula));
                }
            }
        }

        public bool HasSuggestedFormula => !string.IsNullOrWhiteSpace(SuggestedFormula);

        public void NotifyCellsChanged()
        {
            OnPropertyChanged(nameof(HasReferencedCells));
        }
    }
}
