using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using ExcelSupport.Helpers;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.Views
{
    public partial class AiQuickTranslatePopup : Window, INotifyPropertyChanged
    {
        private static AiQuickTranslatePopup? _currentInstance;

        private bool _isDarkTheme;
        private List<CellTextItem> _originalItems = new List<CellTextItem>();
        private List<CellTextItem> _targetItems = new List<CellTextItem>();
        private CancellationTokenSource? _cts;
        private TranslationDirection _currentDirection = TranslationDirection.JapaneseToVietnamese;

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
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public AiQuickTranslatePopup(List<CellTextItem> items, bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            _originalItems = (items ?? new List<CellTextItem>()).Select(x => x.Clone()).ToList();
            _targetItems = _originalItems.Select(x => x.Clone()).ToList();
            DataContext = this;

            Loaded += OnLoaded;
            Closed += (s, e) =>
            {
                _cts?.Cancel();
                if (ReferenceEquals(_currentInstance, this))
                {
                    _currentInstance = null;
                }
            };
        }

        public static void ShowPopup(bool isDarkTheme = false)
        {
            try
            {
                var addIn = AddInEvents.Instance;
                if (addIn == null)
                {
                    WpfMessageBox.Show("Không thể kết nối với Excel.", "Dịch Nhanh AI", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    return;
                }

                var items = addIn.GetSelectedCellsText(maxCells: 10, visibleOnly: true);
                if (items == null || items.Count == 0)
                {
                    var activeInfo = addIn.GetActiveCellInfo();
                    if (activeInfo != null && !string.IsNullOrWhiteSpace(activeInfo.Value))
                    {
                        items = new List<CellTextItem>
                        {
                            new CellTextItem
                            {
                                Address = activeInfo.CellAddress,
                                OriginalText = activeInfo.Value
                            }
                        };
                    }
                }

                if (items == null || items.Count == 0)
                {
                    WpfMessageBox.Show("Ô đang chọn không có nội dung văn bản để dịch.", "Dịch Nhanh AI", 
                                       WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                    return;
                }

                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Close();
                }

                _currentInstance = new AiQuickTranslatePopup(items, isDarkTheme);

                try
                {
                    if (addIn.ExcelAppInstance != null)
                    {
                        var helper = new WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                _currentInstance.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(_currentInstance);
                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở popup dịch nhanh:\n{ex.Message}", "Lỗi Dịch Nhanh", 
                                   WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_originalItems.Count > 0)
            {
                if (_originalItems.Count == 1)
                {
                    TxtCellBadge.Text = _originalItems[0].Address;
                    TxtOriginalText.Text = _originalItems[0].OriginalText;
                }
                else
                {
                    TxtCellBadge.Text = $"{_originalItems[0].Address}.. (+{_originalItems.Count - 1})";
                    TxtOriginalText.Text = string.Join("\n", _originalItems.Select(i => $"• [{i.Address}]: {i.OriginalText}"));
                }

                // Detect direction based on original text
                var detected = AiTranslationService.DetectDirection(_originalItems.Select(i => i.OriginalText));
                if (detected == TranslationDirection.VietnameseToJapanese)
                {
                    _currentDirection = TranslationDirection.VietnameseToJapanese;
                    if (RbViJa != null) RbViJa.IsChecked = true;
                }
                else
                {
                    _currentDirection = TranslationDirection.JapaneseToVietnamese;
                    if (RbJaVi != null) RbJaVi.IsChecked = true;
                }
            }
            else
            {
                if (RbJaVi != null) RbJaVi.IsChecked = true;
            }

            StartTranslation();
        }

        private async void StartTranslation()
        {
            if (_originalItems.Count == 0) return;

            _cts?.Cancel();
            _cts = new CancellationTokenSource();

            PanelLoading.Visibility = Visibility.Visible;
            SvTranslation.Visibility = Visibility.Collapsed;
            TxtLoadingStatus.Text = "Đang dịch với AI...";
            BtnInsertCell.IsEnabled = false;

            try
            {
                var config = AiConfigManager.Current;
                var tone = TranslationTone.ITSoftware;

                var itemsToTranslate = _originalItems.Select(x => x.Clone()).ToList();

                var translatedItems = await AiTranslationService.TranslateBatchAsync(
                    itemsToTranslate,
                    _currentDirection,
                    tone,
                    enableGlossary: true,
                    customGlossary: config.Glossary,
                    progress: null,
                    cancellationToken: _cts.Token
                );

                if (translatedItems != null && translatedItems.Count > 0)
                {
                    _targetItems = translatedItems;
                    RenderResult();
                }
                else
                {
                    TxtTranslationResult.Text = "[Không nhận được kết quả dịch]";
                }
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                TxtTranslationResult.Text = $"[Lỗi dịch: {ex.Message}]";
            }
            finally
            {
                PanelLoading.Visibility = Visibility.Collapsed;
                SvTranslation.Visibility = Visibility.Visible;
                BtnInsertCell.IsEnabled = true;
            }
        }

        private void RenderResult()
        {
            TxtTranslationResult.Inlines.Clear();

            if (_targetItems.Count == 1)
            {
                var item = _targetItems[0];
                if (item.TranslatedRuns != null && item.TranslatedRuns.Count > 0)
                {
                    foreach (var runModel in item.TranslatedRuns)
                    {
                        if (string.IsNullOrEmpty(runModel.Text)) continue;
                        var run = new Run(runModel.Text);
                        if (runModel.IsStrikethrough) run.TextDecorations = TextDecorations.Strikethrough;
                        if (runModel.IsBold) run.FontWeight = FontWeights.Bold;
                        if (runModel.IsItalic) run.FontStyle = FontStyles.Italic;
                        if (!string.IsNullOrEmpty(runModel.ColorHex))
                        {
                            try
                            {
                                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(runModel.ColorHex);
                                run.Foreground = new SolidColorBrush(color);
                            }
                            catch { }
                        }
                        TxtTranslationResult.Inlines.Add(run);
                    }
                }
                else
                {
                    TxtTranslationResult.Text = item.TranslatedText;
                }
            }
            else
            {
                for (int i = 0; i < _targetItems.Count; i++)
                {
                    var item = _targetItems[i];
                    var headerRun = new Run($"[{item.Address}]: ") { FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(37, 99, 235)) };
                    TxtTranslationResult.Inlines.Add(headerRun);

                    if (item.TranslatedRuns != null && item.TranslatedRuns.Count > 0)
                    {
                        foreach (var runModel in item.TranslatedRuns)
                        {
                            if (string.IsNullOrEmpty(runModel.Text)) continue;
                            var run = new Run(runModel.Text);
                            if (runModel.IsStrikethrough) run.TextDecorations = TextDecorations.Strikethrough;
                            if (runModel.IsBold) run.FontWeight = FontWeights.Bold;
                            if (runModel.IsItalic) run.FontStyle = FontStyles.Italic;
                            if (!string.IsNullOrEmpty(runModel.ColorHex))
                            {
                                try
                                {
                                    var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(runModel.ColorHex);
                                    run.Foreground = new SolidColorBrush(color);
                                }
                                catch { }
                            }
                            TxtTranslationResult.Inlines.Add(run);
                        }
                    }
                    else
                    {
                        TxtTranslationResult.Inlines.Add(new Run(item.TranslatedText ?? string.Empty));
                    }

                    if (i < _targetItems.Count - 1)
                    {
                        TxtTranslationResult.Inlines.Add(new LineBreak());
                    }
                }
            }
        }

        private void OnDirectionChanged(object sender, RoutedEventArgs e)
        {
            if (RbJaVi?.IsChecked == true)
            {
                _currentDirection = TranslationDirection.JapaneseToVietnamese;
            }
            else if (RbViJa?.IsChecked == true)
            {
                _currentDirection = TranslationDirection.VietnameseToJapanese;
            }
            else if (RbEnVi?.IsChecked == true)
            {
                _currentDirection = TranslationDirection.EnglishToVietnamese;
            }
            else if (RbAuto?.IsChecked == true)
            {
                _currentDirection = TranslationDirection.AutoDetect;
            }

            if (IsLoaded)
            {
                StartTranslation();
            }
        }

        private void OnHeaderMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void OnCopyClick(object sender, RoutedEventArgs e)
        {
            try
            {
                string textToCopy = string.Join("\n", _targetItems.Select(i => i.TranslatedText));
                if (!string.IsNullOrEmpty(textToCopy))
                {
                    System.Windows.Clipboard.SetText(textToCopy);
                    BtnCopy.Content = "✅ Đã chép!";
                    Task.Delay(1500).ContinueWith(_ => Dispatcher.Invoke(() => BtnCopy.Content = "📋 Sao chép"));
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể sao chép: {ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
            }
        }

        private void OnInsertCellClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var addIn = AddInEvents.Instance;
                if (addIn != null && _targetItems.Count > 0)
                {
                    addIn.WriteTranslatedCells(_targetItems, writeToAdjacentColumn: false);
                }
                Close();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi chèn vào bảng tính:\n{ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void OnOpenFullDialogClick(object sender, RoutedEventArgs e)
        {
            Close();
            AiTranslateDialog.ShowWindow(IsDarkTheme);
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnWindowKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Shift) == 0 && (Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                OnInsertCellClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                OnCopyClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.F3)
            {
                OnOpenFullDialogClick(sender, e);
                e.Handled = true;
            }
        }
    }
}
