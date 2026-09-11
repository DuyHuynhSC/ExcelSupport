using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Win32;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfWindow = System.Windows.Window;

namespace ExcelSupport.Views
{
    public partial class AiTranslateDialog : WpfWindow, INotifyPropertyChanged
    {
        private bool _isDarkTheme;
        private readonly ObservableCollection<CellTextItem> _translationItems = new ObservableCollection<CellTextItem>();
        private readonly ObservableCollection<GlossaryItem> _glossaryItems = new ObservableCollection<GlossaryItem>();
        private readonly ICollectionView _translationView;
        private readonly ICollectionView _glossaryView;
        private CancellationTokenSource? _translationCts;
        private static AiTranslateDialog? _currentInstance;

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

        public AiTranslateDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

            _translationView = CollectionViewSource.GetDefaultView(_translationItems);
            _translationView.Filter = FilterTranslationItem;
            DgTranslations.ItemsSource = _translationView;

            LoadGlossaryFromConfig();
            _glossaryView = CollectionViewSource.GetDefaultView(_glossaryItems);
            _glossaryView.Filter = FilterGlossaryItem;
            DgGlossary.ItemsSource = _glossaryView;

            Loaded += AiTranslateDialog_Loaded;
            Closing += AiTranslateDialog_Closing;
        }

        public static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new AiTranslateDialog(isDarkTheme);

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new System.Windows.Interop.WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(_currentInstance);
                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở Cửa Sổ Dịch Thuật AI:\n{ex.Message}", "Lỗi Giao Diện",
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void AiTranslateDialog_Loaded(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ScanSelectedCells();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void AiTranslateDialog_Closing(object? sender, CancelEventArgs e)
        {
            _translationCts?.Cancel();
        }

        #region Scanning Selected Cells

        private void ScanSelectedCells()
        {
            _translationItems.Clear();

            var addIn = AddInEvents.Instance;
            if (addIn == null)
            {
                TxtSelectionInfo.Text = LocalizationService.Get("AiTranslate_NoConnection");
                return;
            }

            var items = addIn.GetSelectedCellsText(maxCells: 1000, visibleOnly: true);
            foreach (var item in items)
            {
                _translationItems.Add(item);
            }

            TxtSelectionInfo.Text = LocalizationService.Get("AiTranslate_SelectionCountFormat", _translationItems.Count);

            if (_translationItems.Count == 0)
            {
                TxtSelectionInfo.Text = LocalizationService.Get("AiTranslate_NoSelection");
            }
        }

        private void OnRescanSelectionClick(object sender, RoutedEventArgs e)
        {
            ScanSelectedCells();
            _translationView.Refresh();
        }

        #endregion

        #region Translation Execution

        private async void OnTranslateNowClick(object sender, RoutedEventArgs e)
        {
            if (_translationItems.Count == 0)
            {
                WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgNoSelection"), 
                                   LocalizationService.Get("AiTranslate_HeaderTitle"),
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            // Read options
            var direction = GetSelectedDirection();
            var tone = GetSelectedTone();
            bool enableGlossary = ChkEnableGlossary.IsChecked == true;

            // Prepare UI for translation
            BtnTranslateNow.IsEnabled = false;
            BtnInsertToExcel.IsEnabled = false;
            PanelProgress.Visibility = Visibility.Visible;
            PbTranslate.IsIndeterminate = false;
            PbTranslate.Minimum = 0;
            PbTranslate.Maximum = _translationItems.Count;
            PbTranslate.Value = 0;
            TxtProgressStatus.Text = LocalizationService.Get("AiTranslate_ProgressConnecting");

            _translationCts?.Cancel();
            _translationCts = new CancellationTokenSource();

            var progress = new Progress<(int processed, int total, string status)>(report =>
            {
                PbTranslate.Value = report.processed;
                TxtProgressStatus.Text = report.status;
            });

            var stopwatch = Stopwatch.StartNew();

            try
            {
                var inputList = _translationItems.ToList();
                var translatedList = await AiTranslationService.TranslateBatchAsync(
                    inputList,
                    direction,
                    tone,
                    enableGlossary,
                    _glossaryItems.ToList(),
                    progress,
                    _translationCts.Token
                );

                for (int i = 0; i < translatedList.Count; i++)
                {
                    if (i < _translationItems.Count)
                    {
                        _translationItems[i].TranslatedRuns = translatedList[i].TranslatedRuns;
                        _translationItems[i].TranslatedText = translatedList[i].TranslatedText;
                    }
                }

                stopwatch.Stop();
                TxtProgressStatus.Text = LocalizationService.Get("AiTranslate_CompleteFormat", _translationItems.Count, stopwatch.Elapsed.TotalSeconds);
            }
            catch (OperationCanceledException)
            {
                TxtProgressStatus.Text = LocalizationService.Get("AiTranslate_Stopped");
            }
            catch (Exception ex)
            {
                TxtProgressStatus.Text = $"Lỗi: {ex.Message}";
                WpfMessageBox.Show($"Lỗi trong quá trình dịch thuật AI:\n{ex.Message}", "Lỗi Dịch Thuật",
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
            }
            finally
            {
                BtnTranslateNow.IsEnabled = true;
                BtnInsertToExcel.IsEnabled = true;
                _translationView.Refresh();
            }
        }

        private TranslationDirection GetSelectedDirection()
        {
            if (CboDirection.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                if (Enum.TryParse<TranslationDirection>(tag, out var dir))
                {
                    return dir;
                }
            }
            return TranslationDirection.JapaneseToVietnamese;
        }

        private TranslationTone GetSelectedTone()
        {
            if (CboTone.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                if (Enum.TryParse<TranslationTone>(tag, out var tone))
                {
                    return tone;
                }
            }
            return TranslationTone.ITSoftware;
        }

        private void OnChkOverwriteChanged(object sender, RoutedEventArgs e)
        {
            // Visual helper
        }

        #endregion

        #region Search & Filtering

        private bool FilterTranslationItem(object obj)
        {
            if (obj is not CellTextItem item) return false;

            string query = TxtSearch?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(query)) return true;

            return (!string.IsNullOrEmpty(item.OriginalText) && item.OriginalText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (!string.IsNullOrEmpty(item.TranslatedText) && item.TranslatedText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (!string.IsNullOrEmpty(item.Address) && item.Address.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _translationView.Refresh();
        }

        private bool FilterGlossaryItem(object obj)
        {
            if (obj is not GlossaryItem item) return false;

            string query = TxtGlossarySearch?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(query)) return true;

            return (!string.IsNullOrEmpty(item.Japanese) && item.Japanese.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (!string.IsNullOrEmpty(item.Vietnamese) && item.Vietnamese.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (!string.IsNullOrEmpty(item.Note) && item.Note.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void OnGlossarySearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _glossaryView.Refresh();
        }

        private void OnCopyTableClick(object sender, RoutedEventArgs e)
        {
            if (_translationItems.Count == 0) return;

            var sb = new StringBuilder();
            sb.AppendLine($"{LocalizationService.Get("AiTranslate_ColAddress")}\t{LocalizationService.Get("AiTranslate_ColOriginal")}\t{LocalizationService.Get("AiTranslate_ColTranslated")}");
            foreach (var item in _translationItems)
            {
                sb.AppendLine($"{item.Address}\t{item.OriginalText}\t{item.TranslatedText}");
            }

            try
            {
                System.Windows.Clipboard.SetText(sb.ToString());
                WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgCopySuccess"), 
                                   LocalizationService.Get("AiTranslate_HeaderTitle"),
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể sao chép: {ex.Message}", "Lỗi",
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
            }
        }

        #endregion

        #region Glossary Management

        private void LoadGlossaryFromConfig()
        {
            _glossaryItems.Clear();
            var config = AiConfigManager.Current;
            if (config.Glossary != null)
            {
                foreach (var item in config.Glossary)
                {
                    _glossaryItems.Add(item.Clone());
                }
            }
        }

        private void OnAddGlossaryTermClick(object sender, RoutedEventArgs e)
        {
            _glossaryItems.Add(new GlossaryItem
            {
                Japanese = "",
                Vietnamese = "",
                Note = ""
            });
            TabMain.SelectedIndex = 1; // Switch to Glossary tab
        }

        private void OnDeleteGlossaryTermClick(object sender, RoutedEventArgs e)
        {
            if (DgGlossary.SelectedItem is GlossaryItem selected)
            {
                _glossaryItems.Remove(selected);
            }
            else
            {
                WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgSelectDeleteTerm"), 
                                   LocalizationService.Get("AiTranslate_HeaderTitle"),
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
        }

        private void OnSaveGlossaryClick(object sender, RoutedEventArgs e)
        {
            var config = AiConfigManager.Current;
            config.Glossary = _glossaryItems.Where(g => !string.IsNullOrWhiteSpace(g.Japanese) || !string.IsNullOrWhiteSpace(g.Vietnamese)).ToList();
            AiConfigManager.Save(config);

            WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgSaveGlossarySuccess", config.Glossary.Count), 
                               LocalizationService.Get("AiTranslate_HeaderTitle"),
                               WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
        }

        private void OnImportGlossaryCsvClick(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "File CSV/JSON (*.csv;*.json)|*.csv;*.json|Tất cả file (*.*)|*.*",
                Title = LocalizationService.Get("AiTranslate_BtnImportCsv")
            };

            if (ofd.ShowDialog() == true)
            {
                try
                {
                    List<GlossaryItem> imported;
                    if (Path.GetExtension(ofd.FileName).Equals(".json", StringComparison.OrdinalIgnoreCase))
                    {
                        imported = GlossaryService.ImportFromJson(ofd.FileName);
                    }
                    else
                    {
                        imported = GlossaryService.ImportFromCsv(ofd.FileName);
                    }

                    int addedCount = 0;
                    foreach (var item in imported)
                    {
                        if (!_glossaryItems.Any(g => string.Equals(g.Japanese, item.Japanese, StringComparison.OrdinalIgnoreCase)))
                        {
                            _glossaryItems.Add(item);
                            addedCount++;
                        }
                    }

                    WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgImportSuccess", addedCount), 
                                       LocalizationService.Get("AiTranslate_HeaderTitle"),
                                       WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Lỗi nhập file Glossary:\n{ex.Message}", "Lỗi Nhập File",
                                       WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                }
            }
        }

        private void OnExportGlossaryCsvClick(object sender, RoutedEventArgs e)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "File CSV (*.csv)|*.csv|File JSON (*.json)|*.json",
                FileName = $"Glossary_ExcelSupport_{DateTime.Now:yyyyMMdd}.csv",
                Title = LocalizationService.Get("AiTranslate_BtnExportCsv")
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    if (Path.GetExtension(sfd.FileName).Equals(".json", StringComparison.OrdinalIgnoreCase))
                    {
                        GlossaryService.ExportToJson(sfd.FileName, _glossaryItems);
                    }
                    else
                    {
                        GlossaryService.ExportToCsv(sfd.FileName, _glossaryItems);
                    }

                    WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgExportSuccess"), 
                                       LocalizationService.Get("AiTranslate_HeaderTitle"),
                                       WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Lỗi xuất file Glossary:\n{ex.Message}", "Lỗi Xuất File",
                                       WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                }
            }
        }

        #endregion

        #region Dialog Actions (Insert / Cancel)

        private void OnInsertToExcelClick(object sender, RoutedEventArgs e)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null)
            {
                WpfMessageBox.Show(LocalizationService.Get("AiTranslate_NoConnection"), "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                return;
            }

            var itemsToWrite = _translationItems.Where(i => !string.IsNullOrEmpty(i.TranslatedText)).ToList();
            if (itemsToWrite.Count == 0)
            {
                WpfMessageBox.Show(LocalizationService.Get("AiTranslate_MsgNoTransToWrite"), 
                                   LocalizationService.Get("AiTranslate_HeaderTitle"),
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                return;
            }

            bool writeToAdjacent = ChkOverwrite.IsChecked != true;
            bool success = addIn.WriteTranslatedCells(itemsToWrite, writeToAdjacent);

            if (success)
            {
                Close();
            }
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                Close();
            }
        }

        #endregion
    }
}
