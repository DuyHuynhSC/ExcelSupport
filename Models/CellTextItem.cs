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
        private List<TextRunModel>? _formattedRuns;
        public List<TextRunModel>? FormattedRuns
        {
            get => _formattedRuns;
            set
            {
                if (_formattedRuns != value)
                {
                    _formattedRuns = value;
                    OnPropertyChanged(nameof(FormattedRuns));
                }
            }
        }

        private List<TextRunModel>? _translatedRuns;
        public List<TextRunModel>? TranslatedRuns
        {
            get => _translatedRuns;
            set
            {
                if (_translatedRuns != value)
                {
                    _translatedRuns = value;
                    OnPropertyChanged(nameof(TranslatedRuns));
                }
            }
        }

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

        public string GetTaggedOriginalText()
        {
            if (FormattedRuns == null || FormattedRuns.Count == 0)
            {
                return OriginalText;
            }

            var sb = new System.Text.StringBuilder();
            foreach (var run in FormattedRuns)
            {
                if (string.IsNullOrEmpty(run.Text)) continue;

                string openTag = string.Empty;
                string closeTag = string.Empty;

                if (run.IsStrikethrough && !string.IsNullOrEmpty(run.ColorHex))
                {
                    openTag = $"<s color=\"{run.ColorHex}\">";
                    closeTag = "</s>";
                }
                else if (run.IsStrikethrough)
                {
                    openTag = "<s>";
                    closeTag = "</s>";
                }
                else if (!string.IsNullOrEmpty(run.ColorHex))
                {
                    openTag = $"<color hex=\"{run.ColorHex}\">";
                    closeTag = "</color>";
                }

                if (run.IsBold)
                {
                    openTag = "<b>" + openTag;
                    closeTag = closeTag + "</b>";
                }
                if (run.IsItalic)
                {
                    openTag = "<i>" + openTag;
                    closeTag = closeTag + "</i>";
                }

                sb.Append(openTag + run.Text + closeTag);
            }
            return sb.ToString();
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
                TranslatedRuns = TranslatedRuns != null ? new List<TextRunModel>(TranslatedRuns) : null,
                TranslatedText = TranslatedText
            };
        }
    }
}
