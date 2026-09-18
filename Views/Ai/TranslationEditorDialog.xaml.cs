using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelSupport.Models;
using ExcelSupport.Services;

namespace ExcelSupport.Views
{
    public partial class TranslationEditorDialog : Window, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private readonly ObservableCollection<CellTextItem> _items = new ObservableCollection<CellTextItem>();
        private readonly ICollectionView _view;
        private readonly List<CellTextItem> _sourceItems;

        public bool IsDarkTheme
        {
            get => _isDarkTheme;
            set
            {
                if (_isDarkTheme != value)
                {
                    _isDarkTheme = value;
                    OnPropertyChanged(nameof(IsDarkTheme));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public TranslationEditorDialog(List<CellTextItem> items, bool isJaToVi, bool writeToAdjacentColumn, bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            _sourceItems = items ?? new List<CellTextItem>();

            foreach (var item in _sourceItems)
            {
                _items.Add(item);
            }

            _view = CollectionViewSource.GetDefaultView(_items);
            _view.Filter = FilterItem;
            GridTranslations.ItemsSource = _view;

            // Setup Badges
            TxtDirectionBadge.Text = isJaToVi ? LocalizationService.Get("Ai_BtnJaToVi") : LocalizationService.Get("Ai_BtnViToJa");
            TxtCountBadge.Text = string.Format(LocalizationService.Get("TransEdit_CountBadgeFormat"), _sourceItems.Count);
            TxtTargetColBadge.Text = writeToAdjacentColumn ? LocalizationService.Get("TransEdit_TargetAdjacent") : LocalizationService.Get("TransEdit_TargetOverwrite");
        }

        private bool FilterItem(object obj)
        {
            if (obj is not CellTextItem item) return false;

            string query = TxtSearch?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(query)) return true;

            return (item.OriginalText != null && item.OriginalText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (item.TranslatedText != null && item.TranslatedText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (item.Address != null && item.Address.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _view?.Refresh();
        }

        private void OnInsertClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
