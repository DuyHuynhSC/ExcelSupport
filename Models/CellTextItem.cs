using System.ComponentModel;

namespace ExcelSupport.Models
{
    public class CellTextItem : INotifyPropertyChanged
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; } = string.Empty;
        public string OriginalText { get; set; } = string.Empty;

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
                TranslatedText = TranslatedText
            };
        }
    }
}
