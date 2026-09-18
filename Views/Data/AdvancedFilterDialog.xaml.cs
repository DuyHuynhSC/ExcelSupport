using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfBorder = System.Windows.Controls.Border;
using WpfButton = System.Windows.Controls.Button;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfHorizontalAlignment = System.Windows.HorizontalAlignment;
using WpfMessageBox = System.Windows.MessageBox;
using WpfStyle = System.Windows.Style;
using WpfTabControl = System.Windows.Controls.TabControl;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace ExcelSupport.Views
{
    public partial class AdvancedFilterDialog : System.Windows.Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(AdvancedFilterDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static AdvancedFilterDialog? _currentInstance;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new AdvancedFilterDialog(app)
                {
                    IsDarkTheme = isDarkTheme
                };

                try
                {
                    if (app != null)
                    {
                        new System.Windows.Interop.WindowInteropHelper(_currentInstance).Owner = (IntPtr)app.Hwnd;
                    }
                }
                catch { }

                _currentInstance.ShowDialog();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở màn hình Bộ Lọc Nâng Cao:\n{ex.Message}\n\nChi tiết:\n{ex.StackTrace}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static AdvancedFilterSavedState _savedState = new AdvancedFilterSavedState();

        private readonly ExcelApp? _excelApp;
        private Workbook? _targetWb;
        private _Worksheet? _targetWs;
        private List<ColumnHeaderItem> _columns = new List<ColumnHeaderItem>();

        private AdvancedFilterCriteria _visualCriteria = new AdvancedFilterCriteria();
        private BatchListFilterCriteria _batchCriteria = new BatchListFilterCriteria();
        private object[,]? _cachedValues2D;
        private int _cachedStartRow = 1;
        private int _cachedStartCol = 1;
        private bool _isInitialized;

        public AdvancedFilterDialog(ExcelApp? app)
        {
            InitializeComponent();
            _excelApp = app;

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? false;
            }
            catch { }

            Loaded += OnDialogLoaded;
            Closing += (s, e) => SaveCurrentDialogState();
        }

        private void OnDialogLoaded(object sender, RoutedEventArgs e)
        {
            InitializeSheetData();
        }

        private void SaveCurrentDialogState()
        {
            try
            {
                _savedState.SelectedTabIndex = MainTabControl.SelectedIndex;

                // Tab 1
                _savedState.BatchRawText = TxtBatchPaste.Text;
                if (CboBatchColumn.SelectedItem is ColumnHeaderItem bCol)
                {
                    _savedState.BatchTargetColumnIndex = bCol.ColumnIndex;
                    _savedState.BatchTargetColumnName = bCol.HeaderText;
                }
                _savedState.BatchExcludeList = (RbBlacklist.IsChecked == true);
                _savedState.BatchIsExactMatch = (ChkBatchExact.IsChecked == true);
                _savedState.BatchMatchCase = (ChkBatchCase.IsChecked == true);

                // Tab 2
                _savedState.VisualCriteria = _visualCriteria;

                // Tab 3
                _savedState.QuickExpressionText = TxtQuickExpression.Text;
                if (CboQuickColumn.SelectedItem is ColumnHeaderItem qCol)
                {
                    _savedState.QuickTargetColumnIndex = qCol.ColumnIndex;
                    _savedState.QuickTargetColumnName = qCol.HeaderText;
                }
            }
            catch { }
        }

        private void InitializeSheetData()
        {
            if (_excelApp == null) return;

            try
            {
                _isInitialized = false;
                _targetWb = _excelApp.ActiveWorkbook;
                _targetWs = _excelApp.ActiveSheet as _Worksheet;

                if (_targetWs == null)
                {
                    WpfMessageBox.Show("Vui lòng mở một bảng tính Excel để sử dụng tính năng lọc dữ liệu.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Đọc danh sách cột
                _columns = AdvancedFilterService.GetSheetColumns(_targetWs);
                if (_columns.Count == 0) return;

                // Nạp vào ComboBox Tab 1 & Tab 3
                CboBatchColumn.ItemsSource = _columns;
                CboQuickColumn.ItemsSource = _columns;

                // Khôi phục Tab đã chọn trước đó
                if (_savedState.SelectedTabIndex >= 0 && _savedState.SelectedTabIndex < MainTabControl.Items.Count)
                {
                    MainTabControl.SelectedIndex = _savedState.SelectedTabIndex;
                }

                // Khôi phục giá trị đã lọc trước đó cho Tab 1 (Batch List)
                if (!string.IsNullOrEmpty(_savedState.BatchRawText))
                {
                    TxtBatchPaste.Text = _savedState.BatchRawText;
                    _batchCriteria.RawPasteText = _savedState.BatchRawText;
                    _batchCriteria.ParsedItems = AdvancedFilterService.ParseBatchList(_savedState.BatchRawText);
                    TxtBatchCountBadge.Text = LocalizationService.Get("Filter_BatchCountBadge", _batchCriteria.ParsedItems.Count);
                }
                _batchCriteria.ExcludeList = _savedState.BatchExcludeList;
                _batchCriteria.IsExactMatch = _savedState.BatchIsExactMatch;
                _batchCriteria.MatchCase = _savedState.BatchMatchCase;
                RbWhitelist.IsChecked = !_savedState.BatchExcludeList;
                RbBlacklist.IsChecked = _savedState.BatchExcludeList;
                ChkBatchExact.IsChecked = _savedState.BatchIsExactMatch;
                ChkBatchCase.IsChecked = _savedState.BatchMatchCase;

                int batchColIdx = _columns.FindIndex(c => c.ColumnIndex == _savedState.BatchTargetColumnIndex || (!string.IsNullOrEmpty(_savedState.BatchTargetColumnName) && c.HeaderText == _savedState.BatchTargetColumnName));
                CboBatchColumn.SelectedIndex = batchColIdx >= 0 ? batchColIdx : 0;
                if (CboBatchColumn.SelectedItem is ColumnHeaderItem curBCol)
                {
                    _batchCriteria.TargetColumnIndex = curBCol.ColumnIndex;
                    _batchCriteria.TargetColumnName = curBCol.HeaderText;
                }

                // Khôi phục Tab 2 (Visual Builder)
                if (_savedState.VisualCriteria != null && _savedState.VisualCriteria.Groups.Count > 0)
                {
                    _visualCriteria = _savedState.VisualCriteria;
                    RenderVisualGroupsUI();
                }
                else
                {
                    InitDefaultVisualGroups();
                }

                // Khôi phục Tab 3 (Quick Expression)
                if (!string.IsNullOrEmpty(_savedState.QuickExpressionText))
                {
                    TxtQuickExpression.Text = _savedState.QuickExpressionText;
                }
                int quickColIdx = _columns.FindIndex(c => c.ColumnIndex == _savedState.QuickTargetColumnIndex || (!string.IsNullOrEmpty(_savedState.QuickTargetColumnName) && c.HeaderText == _savedState.QuickTargetColumnName));
                CboQuickColumn.SelectedIndex = quickColIdx >= 0 ? quickColIdx : 0;

                // Đọc mảng 2D vào bộ nhớ đệm phục vụ Live Preview siêu tốc
                CacheSheetData();

                _isInitialized = true;

                // Cập nhật xem trước ban đầu
                UpdateLivePreview();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InitializeSheetData error: {ex.Message}");
            }
        }

        private void CacheSheetData()
        {
            if (_targetWs == null) return;

            Range? usedRange = null;
            try
            {
                usedRange = _targetWs.UsedRange;
                if (usedRange != null && usedRange.Rows.Count > 1)
                {
                    _cachedValues2D = (object[,])usedRange.Value2;
                    _cachedStartRow = usedRange.Row;
                    _cachedStartCol = usedRange.Column;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CacheSheetData error: {ex.Message}");
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        #region TAB 1: Batch List Filter

        private void OnBatchColumnChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (CboBatchColumn.SelectedItem is ColumnHeaderItem col)
            {
                _batchCriteria.TargetColumnIndex = col.ColumnIndex;
                _batchCriteria.TargetColumnName = col.HeaderText;
                UpdateLivePreview();
            }
        }

        private void OnBatchPasteTextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            string raw = TxtBatchPaste.Text;
            _batchCriteria.RawPasteText = raw;
            _batchCriteria.ParsedItems = AdvancedFilterService.ParseBatchList(raw);

            TxtBatchCountBadge.Text = LocalizationService.Get("Filter_BatchCountBadge", _batchCriteria.ParsedItems.Count);
            UpdateLivePreview();
        }

        private void OnBatchOptionChanged(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _batchCriteria.ExcludeList = (RbBlacklist.IsChecked == true);
            _batchCriteria.IsExactMatch = (ChkBatchExact.IsChecked == true);
            _batchCriteria.MatchCase = (ChkBatchCase.IsChecked == true);

            UpdateLivePreview();
        }

        private void OnPasteFromClipboardClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (System.Windows.Clipboard.ContainsText())
                {
                    string clipText = System.Windows.Clipboard.GetText();
                    TxtBatchPaste.Text = clipText;
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể đọc từ Clipboard:\n{ex.Message}", "Thông Báo",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnClearBatchTextClick(object sender, RoutedEventArgs e)
        {
            TxtBatchPaste.Text = string.Empty;
        }

        #endregion

        #region TAB 2: Visual Rule Builder

        private void InitDefaultVisualGroups()
        {
            _visualCriteria = new AdvancedFilterCriteria
            {
                OuterOperator = LogicalOperator.Or
            };

            int defaultCol = _columns.Count > 0 ? _columns[0].ColumnIndex : 1;
            string defaultName = _columns.Count > 0 ? _columns[0].HeaderText : "";

            // Tạo nhóm 1 mặc định: > 0 và < 50
            var g1 = new FilterRuleGroup { GroupTitle = LocalizationService.Get("Filter_GroupTitle", 1), InnerOperator = LogicalOperator.And };
            g1.Rules.Add(new FilterRule { ColumnIndex = defaultCol, ColumnName = defaultName, Operator = FilterOperator.GreaterThan, Value1 = "0" });
            g1.Rules.Add(new FilterRule { ColumnIndex = defaultCol, ColumnName = defaultName, Operator = FilterOperator.LessThan, Value1 = "50" });

            // Tạo nhóm 2 mặc định: > 250
            var g2 = new FilterRuleGroup { GroupTitle = LocalizationService.Get("Filter_GroupTitle", 2), InnerOperator = LogicalOperator.And };
            g2.Rules.Add(new FilterRule { ColumnIndex = defaultCol, ColumnName = defaultName, Operator = FilterOperator.GreaterThan, Value1 = "250" });

            _visualCriteria.Groups.Add(g1);
            _visualCriteria.Groups.Add(g2);

            RenderVisualGroupsUI();
        }

        private void RenderVisualGroupsUI()
        {
            StackGroups.Children.Clear();

            for (int gIdx = 0; gIdx < _visualCriteria.Groups.Count; gIdx++)
            {
                var group = _visualCriteria.Groups[gIdx];
                int currentGIdx = gIdx;

                // Khung viền nhóm
                var groupBorder = new WpfBorder
                {
                    CornerRadius = new CornerRadius(8),
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(0, 0, 0, 10),
                    Padding = new Thickness(12, 8, 12, 8)
                };

                groupBorder.Background = new SolidColorBrush(IsDarkTheme ? System.Windows.Media.Color.FromRgb(30, 41, 59) : System.Windows.Media.Color.FromRgb(248, 250, 252));
                groupBorder.BorderBrush = new SolidColorBrush(IsDarkTheme ? System.Windows.Media.Color.FromRgb(51, 65, 85) : System.Windows.Media.Color.FromRgb(226, 232, 240));

                var groupStack = new StackPanel();

                // Group Header
                var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });

                var txtTitle = new TextBlock
                {
                    Text = $"📦 " + LocalizationService.Get("Filter_GroupTitle", (gIdx + 1)),
                    FontWeight = FontWeights.Bold,
                    FontSize = 12.5,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 10, 0)
                };
                txtTitle.Foreground = new SolidColorBrush(IsDarkTheme ? System.Windows.Media.Color.FromRgb(248, 250, 252) : System.Windows.Media.Color.FromRgb(15, 23, 42));
                Grid.SetColumn(txtTitle, 0);

                var cboInnerOp = new WpfComboBox
                {
                    Height = 24,
                    Width = 140,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    FontSize = 11.5
                };
                cboInnerOp.Items.Add(LocalizationService.Get("Filter_InnerAnd"));
                cboInnerOp.Items.Add(LocalizationService.Get("Filter_InnerOr"));
                cboInnerOp.SelectedIndex = group.InnerOperator == LogicalOperator.And ? 0 : 1;
                cboInnerOp.SelectionChanged += (s, e) =>
                {
                    group.InnerOperator = cboInnerOp.SelectedIndex == 0 ? LogicalOperator.And : LogicalOperator.Or;
                    UpdateLivePreview();
                };
                Grid.SetColumn(cboInnerOp, 1);

                // Nút Xóa nhóm
                var btnDelGroup = new WpfButton
                {
                    Content = LocalizationService.Get("Filter_BtnDelGroup"),
                    FontSize = 11,
                    Padding = new Thickness(8, 2, 8, 2),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    Style = (WpfStyle)FindResource("FilterSecondaryBtn")
                };
                btnDelGroup.Click += (s, e) =>
                {
                    if (_visualCriteria.Groups.Count > 1)
                    {
                        _visualCriteria.Groups.RemoveAt(currentGIdx);
                        RenderVisualGroupsUI();
                        UpdateLivePreview();
                    }
                };
                Grid.SetColumn(btnDelGroup, 3);

                headerGrid.Children.Add(txtTitle);
                headerGrid.Children.Add(cboInnerOp);
                headerGrid.Children.Add(btnDelGroup);
                groupStack.Children.Add(headerGrid);

                // Danh sách các dòng Rules trong nhóm
                for (int rIdx = 0; rIdx < group.Rules.Count; rIdx++)
                {
                    var rule = group.Rules[rIdx];
                    int currentRIdx = rIdx;

                    var ruleGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) }); // Cột
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) }); // Toán tử
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) }); // Giá trị 1
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) }); // Giá trị 2 (Between)
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    ruleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });  // Xóa rule

                    // Cột Combobox
                    var cboCol = new WpfComboBox
                    {
                        ItemsSource = _columns,
                        Height = 26,
                        Margin = new Thickness(0, 0, 6, 0),
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    var selCol = _columns.FirstOrDefault(c => c.ColumnIndex == rule.ColumnIndex) ?? _columns.FirstOrDefault();
                    cboCol.SelectedItem = selCol;
                    cboCol.SelectionChanged += (s, e) =>
                    {
                        if (cboCol.SelectedItem is ColumnHeaderItem ch)
                        {
                            rule.ColumnIndex = ch.ColumnIndex;
                            rule.ColumnName = ch.HeaderText;
                            UpdateLivePreview();
                        }
                    };
                    Grid.SetColumn(cboCol, 0);

                    // Toán tử Combobox
                    var cboOp = new WpfComboBox
                    {
                        Height = 26,
                        Margin = new Thickness(0, 0, 6, 0),
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    var opList = Enum.GetValues(typeof(FilterOperator)).Cast<FilterOperator>().ToList();
                    foreach (var op in opList)
                    {
                        cboOp.Items.Add(LocalizationService.GetOperatorDescription(op));
                    }
                    cboOp.SelectedIndex = opList.IndexOf(rule.Operator);
                    Grid.SetColumn(cboOp, 1);

                    // Giá trị 1 Textbox
                    var txtVal1 = new WpfTextBox
                    {
                        Text = rule.Value1,
                        Height = 26,
                        Margin = new Thickness(0, 0, 6, 0),
                        Padding = new Thickness(4, 0, 4, 0),
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    txtVal1.TextChanged += (s, e) =>
                    {
                        rule.Value1 = txtVal1.Text;
                        UpdateLivePreview();
                    };
                    Grid.SetColumn(txtVal1, 2);

                    // Giá trị 2 Textbox (chỉ hiện khi Between)
                    var txtVal2 = new WpfTextBox
                    {
                        Text = rule.Value2,
                        Height = 26,
                        Margin = new Thickness(0, 0, 6, 0),
                        Padding = new Thickness(4, 0, 4, 0),
                        VerticalContentAlignment = VerticalAlignment.Center,
                        Visibility = rule.IsBetweenOperator ? Visibility.Visible : Visibility.Collapsed
                    };
                    txtVal2.TextChanged += (s, e) =>
                    {
                        rule.Value2 = txtVal2.Text;
                        UpdateLivePreview();
                    };
                    Grid.SetColumn(txtVal2, 3);

                    cboOp.SelectionChanged += (s, e) =>
                    {
                        if (cboOp.SelectedIndex >= 0 && cboOp.SelectedIndex < opList.Count)
                        {
                            rule.Operator = opList[cboOp.SelectedIndex];
                            txtVal2.Visibility = rule.IsBetweenOperator ? Visibility.Visible : Visibility.Collapsed;
                            txtVal1.IsEnabled = rule.NeedsValue;
                            UpdateLivePreview();
                        }
                    };

                    // Nút Xóa rule
                    var btnDelRule = new WpfButton
                    {
                        Content = "❌",
                        FontSize = 10,
                        Padding = new Thickness(4, 2, 4, 2),
                        Cursor = System.Windows.Input.Cursors.Hand,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 38, 38)),
                        Style = (WpfStyle)FindResource("FilterSecondaryBtn")
                    };
                    btnDelRule.Click += (s, e) =>
                    {
                        if (group.Rules.Count > 1)
                        {
                            group.Rules.RemoveAt(currentRIdx);
                            RenderVisualGroupsUI();
                            UpdateLivePreview();
                        }
                    };
                    Grid.SetColumn(btnDelRule, 5);

                    ruleGrid.Children.Add(cboCol);
                    ruleGrid.Children.Add(cboOp);
                    ruleGrid.Children.Add(txtVal1);
                    ruleGrid.Children.Add(txtVal2);
                    ruleGrid.Children.Add(btnDelRule);
                    groupStack.Children.Add(ruleGrid);
                }

                // Nút Thêm điều kiện trong nhóm
                var btnAddRule = new WpfButton
                {
                    Content = LocalizationService.Get("Filter_BtnAddRule"),
                    HorizontalAlignment = WpfHorizontalAlignment.Left,
                    FontSize = 11,
                    Padding = new Thickness(8, 3, 8, 3),
                    Margin = new Thickness(0, 4, 0, 0),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    Style = (WpfStyle)FindResource("FilterSecondaryBtn")
                };
                btnAddRule.Click += (s, e) =>
                {
                    int defCol = _columns.Count > 0 ? _columns[0].ColumnIndex : 1;
                    string defName = _columns.Count > 0 ? _columns[0].HeaderText : "";
                    group.Rules.Add(new FilterRule { ColumnIndex = defCol, ColumnName = defName, Operator = FilterOperator.GreaterThan, Value1 = "0" });
                    RenderVisualGroupsUI();
                    UpdateLivePreview();
                };
                groupStack.Children.Add(btnAddRule);

                groupBorder.Child = groupStack;
                StackGroups.Children.Add(groupBorder);
            }
        }

        private static string GetEnumDescription(Enum value)
        {
            var field = value.GetType().GetField(value.ToString());
            if (field != null)
            {
                var attr = Attribute.GetCustomAttribute(field, typeof(System.ComponentModel.DescriptionAttribute)) as System.ComponentModel.DescriptionAttribute;
                if (attr != null) return attr.Description;
            }
            return value.ToString();
        }

        private void OnOuterOperatorChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            _visualCriteria.OuterOperator = CboOuterOperator.SelectedIndex == 0 ? LogicalOperator.Or : LogicalOperator.And;
            UpdateLivePreview();
        }

        private void OnAddGroupClick(object sender, RoutedEventArgs e)
        {
            int nextNum = _visualCriteria.Groups.Count + 1;
            int defCol = _columns.Count > 0 ? _columns[0].ColumnIndex : 1;
            string defName = _columns.Count > 0 ? _columns[0].HeaderText : "";

            var newG = new FilterRuleGroup { GroupTitle = LocalizationService.Get("Filter_GroupTitle", nextNum), InnerOperator = LogicalOperator.And };
            newG.Rules.Add(new FilterRule { ColumnIndex = defCol, ColumnName = defName, Operator = FilterOperator.GreaterThan, Value1 = "0" });

            _visualCriteria.Groups.Add(newG);
            RenderVisualGroupsUI();
            UpdateLivePreview();
        }

        #endregion

        #region TAB 3: Quick Expression

        private void OnQuickColumnChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateLivePreview();
        }

        private void OnQuickExpressionChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateLivePreview();
        }

        private void OnSampleExprClick(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.Content is string s)
            {
                TxtQuickExpression.Text = s;
            }
        }

        #endregion

        #region Live Preview Engine

        private void OnTabSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (e.Source is WpfTabControl)
            {
                UpdateLivePreview();
            }
        }

        private Func<int, bool> GetCurrentRowMatcher()
        {
            if (_cachedValues2D == null) return r => true;

            int selectedTab = MainTabControl.SelectedIndex;

            if (selectedTab == 0) // Tab 1: Batch List
            {
                var exactSet = _batchCriteria.IsExactMatch
                    ? new HashSet<string>(_batchCriteria.ParsedItems, _batchCriteria.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase)
                    : null;
                var list = _batchCriteria.ParsedItems;
                int targetCol = _batchCriteria.TargetColumnIndex;

                return r => AdvancedFilterService.EvaluateBatchListRow(_cachedValues2D, r, targetCol, _cachedStartCol, _batchCriteria, exactSet, list);
            }
            else if (selectedTab == 1) // Tab 2: Visual Builder
            {
                return r => AdvancedFilterService.EvaluateRow(_cachedValues2D, r, _cachedStartCol, _visualCriteria);
            }
            else // Tab 3: Quick Expression
            {
                int defCol = (CboQuickColumn.SelectedItem is ColumnHeaderItem ch) ? ch.ColumnIndex : 1;
                string defName = (CboQuickColumn.SelectedItem is ColumnHeaderItem ch2) ? ch2.HeaderText : "";
                var criteria = AdvancedFilterService.ParseQuickExpression(TxtQuickExpression.Text, defCol, defName);

                return r => AdvancedFilterService.EvaluateRow(_cachedValues2D, r, _cachedStartCol, criteria);
            }
        }

        private void UpdateLivePreview()
        {
            if (!_isInitialized || _targetWs == null || _cachedValues2D == null) return;

            try
            {
                var matcher = GetCurrentRowMatcher();
                var (previewDt, totalRows, matchedRows) = AdvancedFilterService.GetPreviewData(_targetWs, matcher, 12);

                GridPreview.ItemsSource = previewDt.DefaultView;

                double pct = totalRows > 0 ? (double)matchedRows / totalRows * 100.0 : 0;
                TxtPreviewBadge.Text = LocalizationService.Get("Filter_PreviewBadge", matchedRows, totalRows, pct);

                if (matchedRows > 0)
                {
                    BdPreviewBadge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 252, 231)); // #DCFCE7
                    TxtPreviewBadge.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(21, 128, 61));  // #15803D
                }
                else
                {
                    BdPreviewBadge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(254, 226, 226)); // #FEE2E2
                    TxtPreviewBadge.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 38, 38));  // #DC2626
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateLivePreview error: {ex.Message}");
            }
        }

        #endregion

        #region Footer Actions

        private void OnApplyFilterClick(object sender, RoutedEventArgs e)
        {
            if (_targetWs == null) return;

            try
            {
                var matcher = GetCurrentRowMatcher();
                var result = AdvancedFilterService.ApplyInPlaceFilter(_targetWs, matcher);

                if (result.Success)
                {
                    WpfMessageBox.Show($"{result.Message}\n\n(Bạn có thể bấm nút 'Xóa Lọc' bất cứ lúc nào để hiện lại tất cả các dòng).",
                                       "Lọc Thành Công", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show(result.Message, "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi áp dụng bộ lọc:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnHighlightClick(object sender, RoutedEventArgs e)
        {
            if (_targetWs == null) return;

            try
            {
                var matcher = GetCurrentRowMatcher();
                var color = System.Drawing.Color.FromArgb(254, 240, 138);
                int count = AdvancedFilterService.HighlightMatchingRows(_targetWs, matcher, color);

                WpfMessageBox.Show($"Đã tô màu Highlight thành công {count:N0} dòng thỏa mãn điều kiện!", "Tô Màu Thành Công",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi tô màu:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnExtractClick(object sender, RoutedEventArgs e)
        {
            if (_targetWb == null || _targetWs == null) return;

            try
            {
                var matcher = GetCurrentRowMatcher();
                bool ok = AdvancedFilterService.ExtractMatchingRowsToNewSheet(_targetWb, _targetWs, matcher, out int extractedCount);

                if (ok)
                {
                    WpfMessageBox.Show($"Đã trích xuất thành công {extractedCount:N0} dòng sang sheet mới!", "Trích Xuất Thành Công",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    WpfMessageBox.Show("Không có dòng nào thỏa mãn điều kiện để trích xuất.", "Thông Báo",
                                       MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi trích xuất:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnClearFilterClick(object sender, RoutedEventArgs e)
        {
            if (_targetWs == null) return;

            bool ok = AdvancedFilterService.ClearFilter(_targetWs);
            if (ok)
            {
                WpfMessageBox.Show("Đã hiện lại toàn bộ các dòng trên bảng tính!", "Đã Xóa Lọc",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                CacheSheetData();
                UpdateLivePreview();
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
    }

    public class AdvancedFilterSavedState
    {
        public int SelectedTabIndex { get; set; } = 0;

        // Tab 1: Batch List
        public string BatchRawText { get; set; } = string.Empty;
        public int BatchTargetColumnIndex { get; set; } = 1;
        public string BatchTargetColumnName { get; set; } = string.Empty;
        public bool BatchExcludeList { get; set; } = false;
        public bool BatchIsExactMatch { get; set; } = true;
        public bool BatchMatchCase { get; set; } = false;

        // Tab 2: Visual Builder
        public AdvancedFilterCriteria? VisualCriteria { get; set; }

        // Tab 3: Quick Expression
        public string QuickExpressionText { get; set; } = string.Empty;
        public int QuickTargetColumnIndex { get; set; } = 1;
        public string QuickTargetColumnName { get; set; } = string.Empty;
    }
}
