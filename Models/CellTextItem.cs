using System.Collections.Generic;
using System.ComponentModel;

namespace ExcelSupport.Models
{
    public class TextRunModel
    {
        public string Text { get; set; } = string.Empty;
        public bool IsStrikethrough { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string? ColorHex { get; set; }
    }

    public class CellTextItem : INotifyPropertyChanged
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; } = string.Empty;
        public string OriginalText { get; set; } = string.Empty;
        public List<TextRunModel>? FormattedRuns { get; set; }

        private string _translatedText = string.Empty;
        public string TranslatedText
        {
            get => _translatedText;
            set
            {
                if (_translatedText != value)
                {
                    _translatedText = value;
                    OnPropertyChanged(nameof(TranslatedText));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public CellTextItem Clone()
        {
            return new CellTextItem
            {
                Row = Row,
                Column = Column,
                Address = Address,
                OriginalText = OriginalText,
                FormattedRuns = FormattedRuns != null ? new List<TextRunModel>(FormattedRuns) : null,
                TranslatedText = TranslatedText
            };
        }
    }
}
