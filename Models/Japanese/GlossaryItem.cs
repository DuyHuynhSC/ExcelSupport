using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSupport.Models
{
    public class GlossaryItem : INotifyPropertyChanged
    {
        private string _japanese = string.Empty;
        private string _vietnamese = string.Empty;
        private string _note = string.Empty;

        public string Japanese
        {
            get => _japanese;
            set
            {
                if (_japanese != value)
                {
                    _japanese = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Vietnamese
        {
            get => _vietnamese;
            set
            {
                if (_vietnamese != value)
                {
                    _vietnamese = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Note
        {
            get => _note;
            set
            {
                if (_note != value)
                {
                    _note = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public GlossaryItem Clone()
        {
            return new GlossaryItem
            {
                Japanese = this.Japanese,
                Vietnamese = this.Vietnamese,
                Note = this.Note
            };
        }
    }
}
