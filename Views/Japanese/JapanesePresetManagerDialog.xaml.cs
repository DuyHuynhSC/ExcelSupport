using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using WpfColor = System.Windows.Media.Color;
using WpfColorConverter = System.Windows.Media.ColorConverter;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfBrush = System.Windows.Media.Brush;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfFontFamily = System.Windows.Media.FontFamily;
using WpfHorizontalAlignment = System.Windows.HorizontalAlignment;
using WpfButton = System.Windows.Controls.Button;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace ExcelSupport.Views
{
    public partial class JapanesePresetManagerDialog : Window, INotifyPropertyChanged
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(JapanesePresetManagerDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static JapanesePresetManagerDialog? _currentInstance;

        private List<JapaneseFormatPreset> _presets = new List<JapaneseFormatPreset>();
        private string? _activePresetId;
        private JapaneseFormatPreset? _selectedPreset;
        private bool _isUpdatingForm = false;

        public JapanesePresetManagerDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;

            Loaded += OnLoaded;
            Closed += (s, e) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                {
                    _currentInstance = null;
                }
            };
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

                _currentInstance = new JapanesePresetManagerDialog(isDarkTheme);

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
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetManagerDialog] ShowWindow error: {ex.Message}");
            }
        }

        private static List<string>? _cachedSystemFonts;

        private static List<string> GetSystemFontNames()
        {
            if (_cachedSystemFonts != null) return _cachedSystemFonts;

            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var font in System.Windows.Media.Fonts.SystemFontFamilies)
                {
                    string name = font.Source;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        fonts.Add(name);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[JapanesePresetManagerDialog] Error getting system fonts: {ex.Message}");
            }

            // Đảm bảo Tahoma và các font chuẩn thông dụng luôn có mặt
            var standardFonts = new[]
            {
                "Tahoma",
                "Meiryo UI",
                "Yu Gothic UI",
                "MS Gothic",
                "MS Mincho",
                "Segoe UI",
                "Arial",
                "Calibri",
                "Consolas",
                "Times New Roman"
            };

            foreach (var sf in standardFonts)
            {
                fonts.Add(sf);
            }

            _cachedSystemFonts = fonts.OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
            return _cachedSystemFonts;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            cboFontFamily.ItemsSource = GetSystemFontNames();
            LoadData();
        }

        private void LoadData()
        {
            var config = JapanesePresetFormatManager.CurrentConfig;
            _activePresetId = config.ActivePresetId;
            _presets = config.Presets.Select(p => p.Clone()).ToList();

            UpdatePresetDefaultFlags();
            RefreshPresetsList();

            var toSelect = _presets.FirstOrDefault(p => p.Id == _activePresetId) ?? _presets.FirstOrDefault();
            if (toSelect != null)
            {
                lstPresets.SelectedItem = toSelect;
            }
        }

        private void UpdatePresetDefaultFlags()
        {
            foreach (var p in _presets)
            {
                p.IsDefault = (p.Id == _activePresetId);
            }
        }

        private void RefreshPresetsList()
        {
            lstPresets.ItemsSource = null;
            lstPresets.ItemsSource = _presets;
        }

        private void OnPresetSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingForm) return;

            _selectedPreset = lstPresets.SelectedItem as JapaneseFormatPreset;
            PopulateFormFromSelected();
        }

        private void PopulateFormFromSelected()
        {
            if (_selectedPreset == null)
            {
                borderRightPanel.IsEnabled = false;
                return;
            }

            borderRightPanel.IsEnabled = true;
            _isUpdatingForm = true;

            try
            {
                // Default Status Banner
                bool isDef = (_selectedPreset.Id == _activePresetId);
                btnSetDefault.IsEnabled = !isDef;
                if (isDef)
                {
                    lblDefaultStatus.Text = LocalizationService.Get("JpPreset_StatusDefault", "Đang là Preset Mặc Định");
                    lblDefaultStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(21, 128, 61));
                    iconDefaultStatus.Text = "🟢";
                }
                else
                {
                    lblDefaultStatus.Text = LocalizationService.Get("JpPreset_StatusCustom", "Mẫu tùy chọn");
                    lblDefaultStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(100, 116, 139));
                    iconDefaultStatus.Text = "⚪";
                }

                // Name & Description
                txtName.Text = _selectedPreset.Name;
                txtDescription.Text = _selectedPreset.Description;

                // Font Family
                SetSelectedFontName(_selectedPreset.FontName);

                // Font Size
                cboFontSize.Text = _selectedPreset.FontSize.ToString("0.#");

                // Styles
                chkBold.IsChecked = _selectedPreset.IsBold;
                chkItalic.IsChecked = _selectedPreset.IsItalic;
                chkUnderline.IsChecked = _selectedPreset.IsUnderline;

                // Font Color
                cpFontColor.ColorHex = _selectedPreset.FontColorHex;

                // Fill
                chkHasFill.IsChecked = _selectedPreset.HasFill;
                cpFillColor.ColorHex = _selectedPreset.FillColorHex;

                // Alignment
                cboHAlign.SelectedIndex = (int)_selectedPreset.HorizontalAlign;
                cboVAlign.SelectedIndex = (int)_selectedPreset.VerticalAlign;
                chkWrapText.IsChecked = _selectedPreset.WrapText;

                // Border
                cboBorderStyle.SelectedIndex = (int)_selectedPreset.BorderStyle;
                cpBorderColor.ColorHex = _selectedPreset.BorderColorHex;

                // Number format
                chkApplyNumberFormat.IsChecked = _selectedPreset.ApplyNumberFormat;
                txtNumberFormat.Text = _selectedPreset.NumberFormat;
            }
            finally
            {
                _isUpdatingForm = false;
            }

            UpdateLivePreview();
        }

        private void SetSelectedFontName(string? fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) fontName = "Tahoma";

            string? matched = null;
            if (cboFontFamily.ItemsSource is IEnumerable<string> list)
            {
                matched = list.FirstOrDefault(f => string.Equals(f, fontName, StringComparison.OrdinalIgnoreCase));
            }

            if (matched != null)
            {
                cboFontFamily.SelectedItem = matched;
            }
            else
            {
                cboFontFamily.Text = fontName!;
            }
        }

        private string GetSelectedFontName()
        {
            string? name = cboFontFamily.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = (cboFontFamily.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
            if (string.IsNullOrWhiteSpace(name))
            {
                name = cboFontFamily.Text?.Trim();
            }
            return string.IsNullOrWhiteSpace(name) ? "Tahoma" : name!;
        }

        private void OnLivePreviewFieldChanged(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingForm) return;
            UpdateLivePreview();
        }

        private void UpdateLivePreview()
        {
            if (pvwCellBorder == null || pvwCellText == null) return;

            try
            {
                // Font Family
                string fontName = GetSelectedFontName();
                pvwCellText.FontFamily = new WpfFontFamily(fontName);

                // Font Size (preview scale down slightly if too large to fit box)
                double fontSize = 10.0;
                if (double.TryParse(cboFontSize.Text, out double parsedSize) && parsedSize > 4 && parsedSize < 72)
                {
                    fontSize = parsedSize;
                }
                pvwCellText.FontSize = Math.Min(fontSize * 1.15, 20.0);

                // Bold / Italic / Underline
                pvwCellText.FontWeight = (chkBold.IsChecked == true) ? FontWeights.Bold : FontWeights.Normal;
                pvwCellText.FontStyle = (chkItalic.IsChecked == true) ? FontStyles.Italic : FontStyles.Normal;
                pvwCellText.TextDecorations = (chkUnderline.IsChecked == true) ? TextDecorations.Underline : null;

                // Font Color
                var fontBrush = ParseBrush(cpFontColor?.ColorHex, WpfColor.FromRgb(15, 23, 42));
                pvwCellText.Foreground = fontBrush;

                // Fill Color
                bool hasFill = (chkHasFill.IsChecked == true);
                if (hasFill)
                {
                    pvwCellBorder.Background = ParseBrush(cpFillColor?.ColorHex, WpfColor.FromRgb(255, 255, 255));
                }
                else
                {
                    pvwCellBorder.Background = WpfBrushes.Transparent;
                }

                // Horizontal Alignment
                switch (cboHAlign.SelectedIndex)
                {
                    case 1: // Left
                        pvwCellText.HorizontalAlignment = WpfHorizontalAlignment.Left;
                        pvwCellText.TextAlignment = TextAlignment.Left;
                        break;
                    case 2: // Center
                        pvwCellText.HorizontalAlignment = WpfHorizontalAlignment.Center;
                        pvwCellText.TextAlignment = TextAlignment.Center;
                        break;
                    case 3: // Right
                        pvwCellText.HorizontalAlignment = WpfHorizontalAlignment.Right;
                        pvwCellText.TextAlignment = TextAlignment.Right;
                        break;
                    default:
                        pvwCellText.HorizontalAlignment = WpfHorizontalAlignment.Left;
                        pvwCellText.TextAlignment = TextAlignment.Left;
                        break;
                }

                // Vertical Alignment
                switch (cboVAlign.SelectedIndex)
                {
                    case 0: // Top
                        pvwCellText.VerticalAlignment = VerticalAlignment.Top;
                        break;
                    case 1: // Center
                        pvwCellText.VerticalAlignment = VerticalAlignment.Center;
                        break;
                    case 2: // Bottom
                        pvwCellText.VerticalAlignment = VerticalAlignment.Bottom;
                        break;
                    default:
                        pvwCellText.VerticalAlignment = VerticalAlignment.Center;
                        break;
                }

                // Wrap text
                pvwCellText.TextWrapping = (chkWrapText.IsChecked == true) ? TextWrapping.Wrap : TextWrapping.NoWrap;

                // Borders
                var borderBrush = ParseBrush(cpBorderColor?.ColorHex, WpfColor.FromRgb(203, 213, 225));
                pvwCellBorder.BorderBrush = borderBrush;

                int borderIndex = cboBorderStyle.SelectedIndex;
                switch (borderIndex)
                {
                    case 0: // None
                        pvwCellBorder.BorderThickness = new Thickness(0);
                        break;
                    case 1: // AllThin
                        pvwCellBorder.BorderThickness = new Thickness(1);
                        break;
                    case 2: // OutlineMedium
                        pvwCellBorder.BorderThickness = new Thickness(2);
                        break;
                    case 3: // HeaderBottomDouble
                        pvwCellBorder.BorderThickness = new Thickness(1, 1, 1, 3);
                        break;
                    case 4: // Dotted
                        pvwCellBorder.BorderThickness = new Thickness(1);
                        break;
                    default:
                        pvwCellBorder.BorderThickness = new Thickness(1);
                        break;
                }
            }
            catch { }
        }

        private WpfBrush ParseBrush(string? hex, WpfColor fallback)
        {
            if (string.IsNullOrWhiteSpace(hex)) return new SolidColorBrush(fallback);

            try
            {
                string clean = hex!.Trim();
                if (!clean.StartsWith("#")) clean = "#" + clean;
                var color = (WpfColor)WpfColorConverter.ConvertFromString(clean);
                return new SolidColorBrush(color);
            }
            catch
            {
                return new SolidColorBrush(fallback);
            }
        }

        private void OnPickerColorChanged(object? sender, string hex)
        {
            if (_isUpdatingForm) return;
            UpdateLivePreview();
        }

        private void OnQuickFontColorClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Background is SolidColorBrush brush)
            {
                cpFontColor.ColorHex = $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                UpdateLivePreview();
            }
        }

        private void OnQuickFillColorClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Background is SolidColorBrush brush)
            {
                chkHasFill.IsChecked = true;
                cpFillColor.ColorHex = $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                UpdateLivePreview();
            }
        }

        private void OnQuickBorderColorClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Background is SolidColorBrush brush)
            {
                cpBorderColor.ColorHex = $"#{brush.Color.R:X2}{brush.Color.G:X2}{brush.Color.B:X2}";
                UpdateLivePreview();
            }
        }

        private void OnSetDefaultClick(object sender, RoutedEventArgs e)
        {
            if (_selectedPreset == null) return;

            _activePresetId = _selectedPreset.Id;
            UpdatePresetDefaultFlags();
            RefreshPresetsList();
            lstPresets.SelectedItem = _selectedPreset;

            PopulateFormFromSelected();
        }

        private void OnAddPresetClick(object sender, RoutedEventArgs e)
        {
            var newPreset = new JapaneseFormatPreset
            {
                Id = Guid.NewGuid().ToString(),
                Name = LocalizationService.Get("JpPreset_NewPresetDefaultName", "Mẫu Mới (Custom Preset)"),
                Description = LocalizationService.Get("JpPreset_NewPresetDefaultDesc", "Mẫu định dạng do người dùng thiết lập"),
                IsBuiltIn = false,
                FontName = "Tahoma",
                FontSize = 10.0,
                FontColorHex = "#0F172A",
                HasFill = true,
                FillColorHex = "#E2E8F0",
                HorizontalAlign = PresetHorizontalAlign.Left,
                VerticalAlign = PresetVerticalAlign.Center,
                BorderStyle = PresetBorderStyle.AllThin,
                BorderColorHex = "#CBD5E1"
            };

            _presets.Add(newPreset);
            RefreshPresetsList();
            lstPresets.SelectedItem = newPreset;
        }

        private void OnClonePresetClick(object sender, RoutedEventArgs e)
        {
            if (_selectedPreset == null) return;

            SaveFormToPreset(_selectedPreset);
            var clone = _selectedPreset.Duplicate();
            _presets.Add(clone);
            RefreshPresetsList();
            lstPresets.SelectedItem = clone;
        }

        private void OnDeletePresetClick(object sender, RoutedEventArgs e)
        {
            if (_selectedPreset == null) return;

            if (_presets.Count <= 1)
            {
                WpfMessageBox.Show(
                    this,
                    LocalizationService.Get("JpPreset_ErrCannotDeleteLast", "Không thể xóa mẫu định dạng duy nhất còn lại."),
                    LocalizationService.Get("Common_Notice", "Thông Báo"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var confirm = WpfMessageBox.Show(
                this,
                LocalizationService.Get("JpPreset_ConfirmDelete", _selectedPreset.Name),
                LocalizationService.Get("Common_Notice", "Thông Báo"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            int index = _presets.IndexOf(_selectedPreset);
            _presets.Remove(_selectedPreset);

            if (_selectedPreset.Id == _activePresetId)
            {
                _activePresetId = _presets.FirstOrDefault()?.Id;
            }

            UpdatePresetDefaultFlags();
            RefreshPresetsList();

            int nextIndex = Math.Min(index, _presets.Count - 1);
            if (nextIndex >= 0 && nextIndex < _presets.Count)
            {
                lstPresets.SelectedItem = _presets[nextIndex];
            }
        }

        private void OnImportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = LocalizationService.Get("JpPreset_ImportDialogTitle", "Nhập Cấu Hình Presets Từ File JSON")
                };

                if (dlg.ShowDialog() == true)
                {
                    bool success = JapanesePresetFormatManager.ImportFromFile(dlg.FileName);
                    if (success)
                    {
                        LoadData();
                        WpfMessageBox.Show(
                            this,
                            LocalizationService.Get("JpPreset_ImportSuccess", "Đã nhập thành công bộ mẫu định dạng!"),
                            LocalizationService.Get("Common_Notice", "Thông Báo"),
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }
                    else
                    {
                        WpfMessageBox.Show(
                            this,
                            LocalizationService.Get("JpPreset_ImportError", "Không thể đọc dữ liệu hoặc file JSON không đúng định dạng preset."),
                            LocalizationService.Get("Common_Notice", "Thông Báo"),
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, ex.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedPreset != null)
                {
                    SaveFormToPreset(_selectedPreset);
                }

                // Lưu tạm thời danh sách hiện tại trước khi export
                JapanesePresetFormatManager.SaveAllPresets(_presets, _activePresetId);

                var dlg = new SaveFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json",
                    FileName = "excel_support_format_presets.json",
                    Title = LocalizationService.Get("JpPreset_ExportDialogTitle", "Xuất Cấu Hình Presets Ra File JSON")
                };

                if (dlg.ShowDialog() == true)
                {
                    JapanesePresetFormatManager.ExportToFile(dlg.FileName);
                    WpfMessageBox.Show(
                        this,
                        LocalizationService.Get("JpPreset_ExportSuccess", "Đã xuất cấu hình ra file thành công:\n{0}", dlg.FileName),
                        LocalizationService.Get("Common_Notice", "Thông Báo"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, ex.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnResetDefaultsClick(object sender, RoutedEventArgs e)
        {
            var confirm = WpfMessageBox.Show(
                this,
                LocalizationService.Get("JpPreset_ConfirmReset", "Bạn có chắc chắn muốn khôi phục danh sách mẫu về 8 mẫu chuẩn ban đầu của hệ thống?"),
                LocalizationService.Get("Common_Notice", "Thông Báo"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            JapanesePresetFormatManager.ResetToDefaults();
            LoadData();

            WpfMessageBox.Show(
                this,
                LocalizationService.Get("JpPreset_ResetSuccess", "Đã khôi phục thành công các mẫu chuẩn hệ thống!"),
                LocalizationService.Get("Common_Notice", "Thông Báo"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void OnSavePresetClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedPreset != null)
                {
                    SaveFormToPreset(_selectedPreset);
                }

                JapanesePresetFormatManager.SaveAllPresets(_presets, _activePresetId);

                WpfMessageBox.Show(
                    this,
                    LocalizationService.Get("JpPreset_SaveSuccess", "Đã lưu thành công tất cả mẫu định dạng!"),
                    LocalizationService.Get("Common_Notice", "Thông Báo"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, ex.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnApplySelectionClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedPreset != null)
                {
                    SaveFormToPreset(_selectedPreset);
                }

                var app = AddInEvents.Instance?.ExcelAppInstance ?? (ExcelApp)ExcelDna.Integration.ExcelDnaUtil.Application;
                var result = JapanesePresetFormatService.ApplyPresetToSelection(app, _selectedPreset);

                if (result.Success)
                {
                    WpfMessageBox.Show(
                        this,
                        result.Message,
                        LocalizationService.Get("Common_Notice", "Thông Báo"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show(
                        this,
                        result.Message,
                        LocalizationService.Get("Common_Notice", "Thông Báo"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(this, ex.Message, LocalizationService.Get("Common_Notice", "Thông Báo"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveFormToPreset(JapaneseFormatPreset preset)
        {
            preset.Name = string.IsNullOrWhiteSpace(txtName.Text) ? "Custom Preset" : txtName.Text.Trim();
            preset.Description = txtDescription.Text?.Trim() ?? string.Empty;

            preset.FontName = GetSelectedFontName();

            if (double.TryParse(cboFontSize.Text, out double size) && size > 0)
            {
                preset.FontSize = size;
            }

            preset.IsBold = (chkBold.IsChecked == true);
            preset.IsItalic = (chkItalic.IsChecked == true);
            preset.IsUnderline = (chkUnderline.IsChecked == true);

            preset.FontColorHex = string.IsNullOrWhiteSpace(cpFontColor.ColorHex) ? "#0F172A" : cpFontColor.ColorHex.Trim();

            preset.HasFill = (chkHasFill.IsChecked == true);
            preset.FillColorHex = string.IsNullOrWhiteSpace(cpFillColor.ColorHex) ? "#FFFFFF" : cpFillColor.ColorHex.Trim();

            preset.HorizontalAlign = (PresetHorizontalAlign)Math.Max(0, cboHAlign.SelectedIndex);
            preset.VerticalAlign = (PresetVerticalAlign)Math.Max(0, cboVAlign.SelectedIndex);
            preset.WrapText = (chkWrapText.IsChecked == true);

            preset.BorderStyle = (PresetBorderStyle)Math.Max(0, cboBorderStyle.SelectedIndex);
            preset.BorderColorHex = string.IsNullOrWhiteSpace(cpBorderColor.ColorHex) ? "#CBD5E1" : cpBorderColor.ColorHex.Trim();

            preset.ApplyNumberFormat = (chkApplyNumberFormat.IsChecked == true);
            preset.NumberFormat = txtNumberFormat.Text?.Trim() ?? string.Empty;
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
