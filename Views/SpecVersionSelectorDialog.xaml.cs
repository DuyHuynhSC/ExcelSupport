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
using WpfListBox = System.Windows.Controls.ListBox;
using WpfTabControl = System.Windows.Controls.TabControl;

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
        private List<SpecSearchResultItem> _allResultsVi;
        private List<SpecSearchResultItem> _allResultsJa;
        private List<SpecSearchResultItem> _filteredResultsVi = new List<SpecSearchResultItem>();
        private List<SpecSearchResultItem> _filteredResultsJa = new List<SpecSearchResultItem>();
        private List<string> _initialKeywords;

        private WpfListBox? ActiveListBox => (tabDocs?.SelectedItem == tabJa) ? lstFilesJa : lstFilesVi;
        private List<SpecSearchResultItem> ActiveFilteredResults => (tabDocs?.SelectedItem == tabJa) ? _filteredResultsJa : _filteredResultsVi;

        public SpecVersionSelectorDialog(
            ProjectProfile profile, 
            SpecDocumentType docType, 
            string keyword, 
            List<SpecSearchResultItem> initialResults,
            bool isReadOnly,
            bool isDarkTheme = false,
            IEnumerable<string>? initialKeywords = null)
        {
            InitializeComponent();
            _profile = profile;
            _docType = docType;
            _keyword = keyword ?? string.Empty;
            _allResults = initialResults ?? new List<SpecSearchResultItem>();
            _allResultsVi = _allResults.Where(r => string.Equals(r.Language, "vi", StringComparison.OrdinalIgnoreCase)).ToList();
            _allResultsJa = _allResults.Where(r => string.Equals(r.Language, "ja", StringComparison.OrdinalIgnoreCase)).ToList();
            _initialKeywords = initialKeywords != null
                ? initialKeywords.Where(k => !string.IsNullOrWhiteSpace(k)).Distinct().ToList()
                : (string.IsNullOrWhiteSpace(keyword) ? new List<string>() : new List<string> { keyword! });

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
            bool isDarkTheme = false,
            IEnumerable<string>? initialKeywords = null)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new SpecVersionSelectorDialog(profile, docType, keyword, initialResults, isReadOnly, isDarkTheme, initialKeywords);

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
            lblProfileName.Text = _profile.Name;

            bool isMultiKeyword = _initialKeywords != null && _initialKeywords.Count > 1;

            if (isMultiKeyword)
            {
                lblKeyword.Text = LocalizationService.Get(
                    "SpecLauncher_MultiKeywordsBadge",
                    _initialKeywords!.Count,
                    _keyword);
                // Khi chọn một vùng (nhiều ô), txtFilter để TRỐNG để không lọc mất danh sách ban đầu!
                txtFilter.Text = string.Empty;
            }
            else
            {
                lblKeyword.Text = string.IsNullOrWhiteSpace(_keyword) 
                    ? LocalizationService.Get("Common_All_Parens") 
                    : _keyword;
                txtFilter.Text = _keyword;
            }

            RefreshList();

            // Nếu Tab Tiếng Việt không có file nhưng Tab Tiếng Nhật có file -> Tự động chuyển sang Tab Tiếng Nhật
            if (_filteredResultsVi.Count == 0 && _filteredResultsJa.Count > 0)
            {
                tabDocs.SelectedItem = tabJa;
            }
            else
            {
                tabDocs.SelectedItem = tabVi;
            }

            var activeList = ActiveListBox;
            if (activeList != null)
            {
                if (isMultiKeyword && activeList.Items.Count > 0)
                {
                    activeList.SelectAll();
                }
                else if (activeList.Items.Count > 0)
                {
                    activeList.SelectedIndex = 0;
                }
            }

            txtFilter.Focus();
            if (!string.IsNullOrEmpty(txtFilter.Text))
            {
                txtFilter.SelectAll();
            }
        }

        private void RefreshList()
        {
            string filterText = txtFilter.Text.Trim();
            if (string.IsNullOrWhiteSpace(filterText))
            {
                _filteredResultsVi = new List<SpecSearchResultItem>(_allResultsVi);
                _filteredResultsJa = new List<SpecSearchResultItem>(_allResultsJa);
            }
            else
            {
                _filteredResultsVi = _allResultsVi
                    .Where(f => f.FileName.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.RelativeDirectory.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.DetectedVersion.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                _filteredResultsJa = _allResultsJa
                    .Where(f => f.FileName.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.RelativeDirectory.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                f.DetectedVersion.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            for (int i = 0; i < _filteredResultsVi.Count; i++)
            {
                _filteredResultsVi[i].ItemIndex = i + 1;
                _filteredResultsVi[i].IsLatest = (_docType != SpecDocumentType.TestSpec && i == 0);
            }

            for (int i = 0; i < _filteredResultsJa.Count; i++)
            {
                _filteredResultsJa[i].ItemIndex = i + 1;
                _filteredResultsJa[i].IsLatest = (_docType != SpecDocumentType.TestSpec && i == 0);
            }

            lstFilesVi.ItemsSource = _filteredResultsVi;
            lstFilesJa.ItemsSource = _filteredResultsJa;

            tabVi.Header = LocalizationService.Get("SpecLauncher_TabViFormat", _filteredResultsVi.Count);
            tabJa.Header = LocalizationService.Get("SpecLauncher_TabJaFormat", _filteredResultsJa.Count);

            UpdateOpenButtonText();
        }

        private void OnTabSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is WpfTabControl)
            {
                UpdateOpenButtonText();
            }
        }

        private void UpdateOpenButtonText()
        {
            if (btnOpenSelected != null && ActiveListBox != null)
            {
                int count = ActiveListBox.SelectedItems.Count;
                if (count > 1)
                {
                    btnOpenSelected.Content = LocalizationService.Get("SpecLauncher_BtnOpenFilesFormat", count);
                }
                else
                {
                    btnOpenSelected.Content = LocalizationService.Get("SpecLauncher_BtnOpenFile");
                }
            }
        }

        private void OnFilterTextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshList();
        }

        private void OnFilterPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var activeList = ActiveListBox;
            if (e.Key == Key.Down && activeList != null && activeList.Items.Count > 0)
            {
                activeList.Focus();
                if (activeList.SelectedIndex < 0) activeList.SelectedIndex = 0;
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

            var activeList = ActiveListBox;
            var activeFiltered = ActiveFilteredResults;

            // Phím tắt chọn nhanh 1 - 9:
            bool isListFocused = activeList != null && (activeList.IsFocused || activeList.IsKeyboardFocusWithin);
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

            if (numberPressed > 0 && activeList != null && (isListFocused || isAltPressed))
            {
                int targetIndex = numberPressed - 1;
                if (targetIndex >= 0 && targetIndex < activeFiltered.Count)
                {
                    activeList.SelectedIndex = targetIndex;
                    OpenSelectedFile();
                    e.Handled = true;
                    return;
                }
            }

            if (e.Key == Key.A && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                activeList?.SelectAll();
                e.Handled = true;
                return;
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
            UpdateOpenButtonText();
        }

        private void OnRescanClick(object sender, RoutedEventArgs e)
        {
            string searchKw = txtFilter.Text.Trim();
            if (!string.IsNullOrWhiteSpace(searchKw))
            {
                _allResults = ProjectDocumentLauncherService.SearchFiles(_profile, _docType, searchKw);
            }
            else if (_initialKeywords != null && _initialKeywords.Count > 0)
            {
                _allResults = ProjectDocumentLauncherService.SearchFiles(_profile, _docType, _initialKeywords);
            }
            else
            {
                _allResults = ProjectDocumentLauncherService.SearchFiles(_profile, _docType, string.Empty);
            }

            _allResultsVi = _allResults.Where(r => string.Equals(r.Language, "vi", StringComparison.OrdinalIgnoreCase)).ToList();
            _allResultsJa = _allResults.Where(r => string.Equals(r.Language, "ja", StringComparison.OrdinalIgnoreCase)).ToList();

            RefreshList();
        }

        private void OnOpenSelectedClick(object sender, RoutedEventArgs e)
        {
            OpenSelectedFile();
        }

        private void OpenSelectedFile()
        {
            var activeList = ActiveListBox;
            var activeFiltered = ActiveFilteredResults;

            if (activeList == null) return;

            var selectedItems = activeList.SelectedItems.Cast<SpecSearchResultItem>().ToList();
            if (selectedItems.Count == 0)
            {
                if (activeFiltered.Count > 0)
                {
                    selectedItems.Add(activeFiltered[0]);
                }
                else
                {
                    System.Windows.MessageBox.Show(
                        this,
                        LocalizationService.Get("SpecLauncher_NoFileAvailable", "Không có file tài liệu nào trong danh sách để mở."),
                        "Thông Báo",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }

            // Kiểm tra các file có tồn tại trên đĩa không trước khi đóng dialog
            var missingFiles = selectedItems.Where(s => !File.Exists(s.FilePath)).ToList();
            if (missingFiles.Count > 0)
            {
                if (missingFiles.Count == 1 && selectedItems.Count == 1)
                {
                    System.Windows.MessageBox.Show(
                        this,
                        LocalizationService.Get("SpecLauncher_FileNotFound", missingFiles[0].FilePath),
                        "Thông Báo",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
                else
                {
                    string missingList = string.Join("\n", missingFiles.Select(m => $"• {m.FileName} ({m.FilePath})"));
                    System.Windows.MessageBox.Show(
                        this,
                        LocalizationService.Get("SpecLauncher_SomeFilesNotFound", missingList),
                        "Thông Báo",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }

            var validFiles = selectedItems.Where(s => File.Exists(s.FilePath)).ToList();
            if (validFiles.Count == 0) return;

            bool isReadOnly = chkReadOnly.IsChecked == true;
            var filePaths = validFiles.Select(f => f.FilePath).ToList();
            Close();

            // Mở an toàn qua QueueAsMacro trên luồng Excel để tránh lỗi không xác định và crash Excel
            ProjectDocumentLauncherService.OpenMultipleFilesAsync(filePaths, isReadOnly);
        }

        private void OnOpenFolderClick(object sender, RoutedEventArgs e)
        {
            OpenSelectedFolder();
        }

        private void OpenSelectedFolder()
        {
            var activeList = ActiveListBox;
            var selected = activeList?.SelectedItem as SpecSearchResultItem;
            if (selected != null)
            {
                if (File.Exists(selected.FilePath))
                {
                    ProjectDocumentLauncherService.OpenContainingFolder(selected.FilePath);
                    return;
                }
                else
                {
                    System.Windows.MessageBox.Show(
                        this,
                        LocalizationService.Get("SpecLauncher_FileNotFound", selected.FilePath),
                        "Thông Báo",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }

            string targetLang = (tabDocs.SelectedItem == tabJa) ? "ja" : "vi";
            string targetFolder = _profile.GetTargetFolder(_docType, targetLang);
            if (Directory.Exists(targetFolder))
            {
                ProjectDocumentLauncherService.OpenContainingFolder(targetFolder);
            }
            else
            {
                System.Windows.MessageBox.Show(
                    this,
                    LocalizationService.Get("SpecProfile_TestFolderNotExistPrompt", targetFolder),
                    "Thông Báo",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void OnOpenSettingsClick(object sender, RoutedEventArgs e)
        {
            Close();
            ProjectProfileSettingsDialog.ShowWindow(IsDarkTheme);
        }
    }
}
