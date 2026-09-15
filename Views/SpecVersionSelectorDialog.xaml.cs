using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public partial class SpecVersionSelectorDialog : Window, INotifyPropertyChanged
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(SpecVersionSelectorDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static SpecVersionSelectorDialog? _currentInstance;

        private ProjectProfile _profile;
        private SpecDocumentType _docType;
        private string _keyword;
        private List<SpecSearchResultItem> _allResults;
        private List<SpecSearchResultItem> _filteredResults;

        public SpecVersionSelectorDialog(
            ProjectProfile profile, 
            SpecDocumentType docType, 
            string keyword, 
            List<SpecSearchResultItem> initialResults,
            bool isReadOnly,
            bool isDarkTheme = false)
        {
            InitializeComponent();
            _profile = profile;
            _docType = docType;
            _keyword = keyword;
            _allResults = initialResults ?? new List<SpecSearchResultItem>();
            _filteredResults = new List<SpecSearchResultItem>(_allResults);

            IsDarkTheme = isDarkTheme;
            chkReadOnly.IsChecked = isReadOnly;

            Loaded += OnLoaded;
            Closed += (s, e) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                {
                    _currentInstance = null;
                }
            };
        }

        public static void ShowWindow(
            ProjectProfile profile,
            SpecDocumentType docType,
            string keyword,
            List<SpecSearchResultItem> initialResults,
            bool isReadOnly = false,
            bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new SpecVersionSelectorDialog(profile, docType, keyword, initialResults, isReadOnly, isDarkTheme);

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(_currentInstance);
                _currentInstance.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SpecVersionSelectorDialog] ShowWindow error: {ex.Message}");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            lblDocType.Text = ProjectDocumentLauncherService.GetDocTypeShortBadge(_docType);
            lblKeyword.Text = string.IsNullOrWhiteSpace(_keyword) 
                ? LocalizationService.Get("Common_All_Parens", "(Tất cả)") 
                : _keyword;
            lblProfileName.Text = _profile.Name;
            txtFilter.Text = _keyword;

            RefreshList();

            if (lstFiles.Items.Count > 0)
            {
                lstFiles.SelectedIndex = 0;
            }

            txtFilter.Focus();
            txtFilter.SelectAll();
        }

        private void RefreshList()
        {
            string filterText = txtFilter.Text.Trim();
            if (string.IsNullOrWhiteSpace(filterText))
            {
                _filteredResults = new List<SpecSearchResultItem>(_allResults);
            }
            else
            {
                _filteredResults = _allResults
                    .Where(f => f.FileName.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.RelativeDirectory.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.DetectedVersion.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            for (int i = 0; i < _filteredResults.Count; i++)
            {
                _filteredResults[i].ItemIndex = i + 1;
                _filteredResults[i].IsLatest = (i == 0);
            }

            lstFiles.ItemsSource = _filteredResults;

            if (_filteredResults.Count > 0)
            {
                lstFiles.SelectedIndex = 0;
            }
        }

        private void OnFilterTextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshList();
        }

        private void OnFilterPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Down && lstFiles.Items.Count > 0)
            {
                lstFiles.Focus();
                if (lstFiles.SelectedIndex < 0) lstFiles.SelectedIndex = 0;
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                OpenSelectedFile();
                e.Handled = true;
            }
        }

        private void OnWindowKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.F5)
            {
                OpenSelectedFolder();
                e.Handled = true;
                return;
            }

            // Phím tắt chọn nhanh 1 - 9:
            // Khi người dùng đang focus ở danh sách hoặc bấm kèm Alt
            bool isListFocused = lstFiles.IsFocused || lstFiles.IsKeyboardFocusWithin;
            bool isAltPressed = (Keyboard.Modifiers & ModifierKeys.Alt) == ModifierKeys.Alt;

            int numberPressed = -1;
            if (e.Key >= Key.D1 && e.Key <= Key.D9)
            {
                numberPressed = e.Key - Key.D1 + 1;
            }
            else if (e.Key >= Key.NumPad1 && e.Key <= Key.NumPad9)
            {
                numberPressed = e.Key - Key.NumPad1 + 1;
            }

            if (numberPressed > 0 && (isListFocused || isAltPressed))
            {
                int targetIndex = numberPressed - 1;
                if (targetIndex >= 0 && targetIndex < _filteredResults.Count)
                {
                    lstFiles.SelectedIndex = targetIndex;
                    OpenSelectedFile();
                    e.Handled = true;
                    return;
                }
            }

            if (e.Key == Key.Enter && isListFocused)
            {
                OpenSelectedFile();
                e.Handled = true;
            }
        }

        private void OnListMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenSelectedFile();
        }

        private void OnListSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Tùy chọn cập nhật giao diện khi chọn item
        }

        private void OnRescanClick(object sender, RoutedEventArgs e)
        {
            string searchKw = txtFilter.Text.Trim();
            _allResults = ProjectDocumentLauncherService.SearchFiles(_profile, _docType, searchKw);
            RefreshList();
        }

        private void OnOpenSelectedClick(object sender, RoutedEventArgs e)
        {
            OpenSelectedFile();
        }

        private void OpenSelectedFile()
        {
            var selected = lstFiles.SelectedItem as SpecSearchResultItem;
            if (selected == null)
            {
                if (_filteredResults.Count > 0)
                {
                    selected = _filteredResults[0];
                }
                else
                {
                    return;
                }
            }

            bool isReadOnly = chkReadOnly.IsChecked == true;
            Close();

            // Mở file sau khi dialog đóng
            ProjectDocumentLauncherService.OpenFile(selected.FilePath, isReadOnly);
        }

        private void OnOpenFolderClick(object sender, RoutedEventArgs e)
        {
            OpenSelectedFolder();
        }

        private void OpenSelectedFolder()
        {
            var selected = lstFiles.SelectedItem as SpecSearchResultItem;
            if (selected != null && File.Exists(selected.FilePath))
            {
                ProjectDocumentLauncherService.OpenContainingFolder(selected.FilePath);
            }
            else
            {
                string targetFolder = _profile.GetTargetFolder(_docType);
                ProjectDocumentLauncherService.OpenContainingFolder(targetFolder);
            }
        }

        private void OnOpenSettingsClick(object sender, RoutedEventArgs e)
        {
            Close();
            ProjectProfileSettingsDialog.ShowWindow(IsDarkTheme);
        }
    }
}
