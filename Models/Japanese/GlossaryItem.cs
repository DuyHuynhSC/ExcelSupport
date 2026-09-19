using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSupport.Models
{
    public class GlossaryItem : ObservableModel
    {
        private string _japanese = string.Empty;
        private string _vietnamese = string.Empty;
        private string _note = string.Empty;

        public string Japanese
        {
            get => _japanese;
            set => SetProperty(ref _japanese, value);
        }

        public string Vietnamese
        {
            get => _vietnamese;
            set => SetProperty(ref _vietnamese, value);
        }

        public string Note
        {
            get => _note;
            set => SetProperty(ref _note, value);
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
