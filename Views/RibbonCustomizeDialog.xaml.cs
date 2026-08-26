using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelSupport.Ribbon;
using ExcelSupport.Services;
using ExcelSupport.ViewModels;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfCheckBox = System.Windows.Controls.CheckBox;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using MediaColor = System.Windows.Media.Color;
using WpfOrientation = System.Windows.Controls.Orientation;

namespace ExcelSupport.Views
{
    public partial class RibbonCustomizeDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(RibbonCustomizeDialog),
                new PropertyMetadata(false, OnIsDarkThemeChanged));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static void OnIsDarkThemeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is RibbonCustomizeDialog dlg && e.NewValue is bool isDark)
            {
                if (dlg.aiSettingsControl?.DataContext is ViewModelBase vm)
                {
                    vm.IsDarkTheme = isDark;
                }
            }
        }

        private static RibbonCustomizeDialog? _currentInstance;
        private readonly List<RibbonControlMetadata> _allControls;
        private readonly Dictionary<string, WpfCheckBox> _checkBoxMap = new();
        private readonly Dictionary<string, Border> _cardMap = new();

        public RibbonCustomizeDialog(int initialTabIndex = 0)
        {
            InitializeComponent();
            _allControls = RibbonVisibilityService.GetAllControlsMetadata();
            aiSettingsControl.DataContext = AddInEvents.MainViewModel?.AiSettings ?? new AiSettingsViewModel();
            InitThemeSelection();
            BuildGroupSections();

            if (initialTabIndex > 0 && initialTabIndex < mainTabControl.Items.Count)
            {
                mainTabControl.SelectedIndex = initialTabIndex;
            }

            rbLightTheme.Checked += OnThemeRadioChecked;
            rbDarkTheme.Checked += OnThemeRadioChecked;
        }

        private void OnThemeRadioChecked(object sender, RoutedEventArgs e)
        {
            bool isDark = rbDarkTheme.IsChecked == true;
            IsDarkTheme = isDark;
        }

        private void InitThemeSelection()
        {
            try
            {
                bool isDark = AddInEvents.MainViewModel?.IsDarkTheme ?? AiConfigManager.Current.IsDarkTheme;
                IsDarkTheme = isDark;
                if (isDark)
                {
                    rbDarkTheme.IsChecked = true;
                }
                else
                {
                    rbLightTheme.IsChecked = true;
                }

                if (aiSettingsControl?.DataContext is ViewModelBase vm)
                {
                    vm.IsDarkTheme = isDark;
                }
            }
            catch { }
        }

        public static void ShowWindow(int initialTabIndex = 0)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    if (initialTabIndex >= 0 && initialTabIndex < _currentInstance.mainTabControl.Items.Count)
                    {
                        _currentInstance.mainTabControl.SelectedIndex = initialTabIndex;
                    }
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new RibbonCustomizeDialog(initialTabIndex);

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

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở hộp thoại Cài đặt:\n{ex.Message}",
                                   "Lỗi giao diện", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BuildGroupSections()
        {
            pnlGroups.Children.Clear();
            _checkBoxMap.Clear();
            _cardMap.Clear();

            // Group by GroupId
            var grouped = _allControls.GroupBy(c => c.GroupId).ToList();

            // Prioritize grpDataTools first
            grouped = grouped.OrderBy(g => g.Key == "grpDataTools" ? 0 : 1).ToList();

            foreach (var group in grouped)
            {
                string groupTitle = LocalizationService.Get(group.First().GroupNameKey);
                if (string.IsNullOrWhiteSpace(groupTitle) || groupTitle == group.First().GroupNameKey)
                {
                    groupTitle = group.Key;
                }

                // Group Outer Container
                var groupContainer = new Border
                {
                    Background = new SolidColorBrush(MediaColor.FromRgb(255, 255, 255)),
                    BorderBrush = new SolidColorBrush(MediaColor.FromRgb(226, 232, 240)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(8),
                    Margin = new Thickness(0, 0, 0, 14),
                    Padding = new Thickness(14, 12, 14, 12)
                };

                var groupStack = new StackPanel();

                // Group Header
                var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var headerTitleStack = new StackPanel { Orientation = WpfOrientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                var groupIcon = group.Key == "grpDataTools" ? "📊" :
                                group.Key == "grpAuditTools" ? "🔍" :
                                group.Key == "grpQuickTools" ? "⚡" :
                                group.Key == "grpFileTools" ? "📁" : "🤖";

                headerTitleStack.Children.Add(new TextBlock
                {
                    Text = groupIcon,
                    FontSize = 14,
                    Margin = new Thickness(0, 0, 6, 0),
                    VerticalAlignment = VerticalAlignment.Center
                });

                headerTitleStack.Children.Add(new TextBlock
                {
                    Text = groupTitle,
                    FontWeight = FontWeights.Bold,
                    FontSize = 13,
                    Foreground = new SolidColorBrush(MediaColor.FromRgb(15, 23, 42)),
                    VerticalAlignment = VerticalAlignment.Center
                });

                headerGrid.Children.Add(headerTitleStack);
                Grid.SetColumn(headerTitleStack, 0);

                // Group Toggle All Checkbox
                var chkGroupAll = new WpfCheckBox
                {
                    Content = LocalizationService.Get("RibbonCustomizer_ToggleGroup"),
                    FontSize = 11,
                    Foreground = new SolidColorBrush(MediaColor.FromRgb(100, 116, 139)),
                    VerticalAlignment = VerticalAlignment.Center,
                    IsChecked = group.All(c => c.IsVisible)
                };

                var groupItemsList = group.ToList();
                chkGroupAll.Click += (s, e) =>
                {
                    bool isCheck = chkGroupAll.IsChecked == true;
                    foreach (var item in groupItemsList)
                    {
                        if (_checkBoxMap.TryGetValue(item.ControlId, out var cb))
                        {
                            cb.IsChecked = isCheck;
                        }
                    }
                };

                headerGrid.Children.Add(chkGroupAll);
                Grid.SetColumn(chkGroupAll, 1);

                groupStack.Children.Add(headerGrid);

                // Items WrapGrid
                var itemsGrid = new WrapPanel { Orientation = WpfOrientation.Horizontal };

                foreach (var ctrl in group)
                {
                    string ctrlLabel = LocalizationService.Get(ctrl.NameKey);
                    if (string.IsNullOrWhiteSpace(ctrlLabel) || ctrlLabel == ctrl.NameKey)
                    {
                        ctrlLabel = ctrl.ControlId;
                    }

                    var cardBorder = new Border
                    {
                        Width = 205,
                        Height = 40,
                        Margin = new Thickness(0, 0, 10, 8),
                        Background = new SolidColorBrush(MediaColor.FromRgb(248, 250, 252)),
                        BorderBrush = new SolidColorBrush(MediaColor.FromRgb(226, 232, 240)),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(6),
                        Padding = new Thickness(8, 0, 8, 0),
                        Tag = $"{ctrlLabel} {ctrl.ControlId}".ToLowerInvariant()
                    };

                    var chk = new WpfCheckBox
                    {
                        IsChecked = ctrl.IsVisible,
                        VerticalAlignment = VerticalAlignment.Center,
                        FontSize = 11.5,
                        Tag = ctrl.ControlId
                    };

                    var labelStack = new StackPanel { Orientation = WpfOrientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                    labelStack.Children.Add(new TextBlock
                    {
                        Text = ctrl.IconEmoji,
                        FontSize = 13,
                        Margin = new Thickness(0, 0, 6, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    labelStack.Children.Add(new TextBlock
                    {
                        Text = ctrlLabel,
                        FontSize = 11.5,
                        FontWeight = FontWeights.Medium,
                        Foreground = new SolidColorBrush(MediaColor.FromRgb(30, 41, 59)),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        VerticalAlignment = VerticalAlignment.Center
                    });

                    chk.Content = labelStack;
                    cardBorder.Child = chk;

                    // Hover effect
                    cardBorder.MouseEnter += (s, e) =>
                    {
                        cardBorder.Background = new SolidColorBrush(MediaColor.FromRgb(241, 245, 249));
                    };
                    cardBorder.MouseLeave += (s, e) =>
                    {
                        cardBorder.Background = new SolidColorBrush(MediaColor.FromRgb(248, 250, 252));
                    };

                    _checkBoxMap[ctrl.ControlId] = chk;
                    _cardMap[ctrl.ControlId] = cardBorder;

                    itemsGrid.Children.Add(cardBorder);
                }

                groupStack.Children.Add(itemsGrid);
                groupContainer.Child = groupStack;

                pnlGroups.Children.Add(groupContainer);
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = txtSearch.Text.Trim().ToLowerInvariant();

            foreach (var kvp in _cardMap)
            {
                string tag = kvp.Value.Tag?.ToString() ?? "";
                if (string.IsNullOrEmpty(query) || tag.Contains(query))
                {
                    kvp.Value.Visibility = Visibility.Visible;
                }
                else
                {
                    kvp.Value.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in _checkBoxMap.Values)
            {
                cb.IsChecked = true;
            }
        }

        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cb in _checkBoxMap.Values)
            {
                cb.IsChecked = false;
            }
        }

        private void BtnResetDefault_Click(object sender, RoutedEventArgs e)
        {
            var confirm = WpfMessageBox.Show(
                LocalizationService.Get("RibbonCustomizer_ResetConfirm"),
                LocalizationService.Get("RibbonCustomizer_ResetTitle"),
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                RibbonVisibilityService.ResetToDefault();
                foreach (var cb in _checkBoxMap.Values)
                {
                    cb.IsChecked = true;
                }
                rbLightTheme.IsChecked = true;
                RibbonController.Instance?.InvalidateRibbon();
                txtStatus.Text = LocalizationService.Get("RibbonCustomizer_ResetSuccess");
            }
        }

        private void BtnSaveApply_Click(object sender, RoutedEventArgs e)
        {
            var map = new Dictionary<string, bool>();
            foreach (var kvp in _checkBoxMap)
            {
                map[kvp.Key] = kvp.Value.IsChecked == true;
            }

            bool isDark = rbDarkTheme.IsChecked == true;
            try
            {
                var aiConfig = AiConfigManager.Current;
                aiConfig.IsDarkTheme = isDark;
                AiConfigManager.Save(aiConfig);
                if (AddInEvents.MainViewModel != null)
                {
                    AddInEvents.MainViewModel.IsDarkTheme = isDark;
                }
            }
            catch { }

            bool ok = RibbonVisibilityService.SaveVisibilityMap(map);
            if (ok)
            {
                RibbonController.Instance?.InvalidateRibbon();
                WpfMessageBox.Show(
                    LocalizationService.Get("RibbonCustomizer_SaveSuccess"),
                    LocalizationService.Get("RibbonCustomizer_SaveTitle"),
                    WpfMessageBoxButton.OK,
                    WpfMessageBoxImage.Information);
                Close();
            }
            else
            {
                WpfMessageBox.Show("Có lỗi khi lưu cấu hình Ribbon.", "Lỗi Lưu", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, WpfKeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
            else if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                BtnSaveApply_Click(btnSaveApply, new RoutedEventArgs());
            }
        }
    }
}
