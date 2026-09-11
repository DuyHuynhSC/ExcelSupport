using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfApplication = System.Windows.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using WpfKey = System.Windows.Input.Key;

namespace ExcelSupport.Views
{
    public partial class QuickCommandPaletteDialog : Window, INotifyPropertyChanged
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(QuickCommandPaletteDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static QuickCommandPaletteDialog? _currentInstance;
        private bool _isExecuting = false;

        public QuickCommandPaletteDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;
            DataContext = this;

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
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    _currentInstance.txtSearch.Focus();
                    _currentInstance.txtSearch.SelectAll();
                    return;
                }

                _currentInstance = new QuickCommandPaletteDialog(isDarkTheme);

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

                PaletteCommandItem? commandToRun = null;

                _currentInstance.CommandSelected += (cmd) =>
                {
                    commandToRun = cmd;
                };

                _currentInstance.ShowDialog();
                _currentInstance = null;

                // Thực thi lệnh sau khi popup đã đóng hoàn toàn
                if (commandToRun != null)
                {
                    try
                    {
                        CommandPaletteService.RecordCommandExecution(commandToRun.Id);
                        commandToRun.Action?.Invoke();
                    }
                    catch (Exception ex)
                    {
                        WpfMessageBox.Show($"Lỗi thực thi lệnh '{commandToRun.Title}':\n{ex.Message}",
                                           LocalizationService.Get("btnCommandPalette"),
                                           MessageBoxButton.OK,
                                           MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở Command Palette:\n{ex.Message}", "Command Palette",
                                   MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public event Action<PaletteCommandItem>? CommandSelected;

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ApplyFilter(string.Empty);
            txtSearch.Focus();
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter(txtSearch.Text);
        }

        private void ApplyFilter(string query)
        {
            var results = CommandPaletteService.Search(query);
            lstCommands.ItemsSource = results;

            if (results.Count > 0)
            {
                pnlEmptyState.Visibility = Visibility.Collapsed;
                lstCommands.Visibility = Visibility.Visible;
                lstCommands.SelectedIndex = 0;
            }
            else
            {
                pnlEmptyState.Visibility = Visibility.Visible;
                lstCommands.Visibility = Visibility.Collapsed;
            }

            txtResultCount.Text = $"{results.Count} {LocalizationService.Get("Palette_AvailableCount")}";
        }

        private void OnSearchPreviewKeyDown(object sender, WpfKeyEventArgs e)
        {
            if (e.Key == WpfKey.Down)
            {
                if (lstCommands.SelectedIndex < lstCommands.Items.Count - 1)
                {
                    lstCommands.SelectedIndex++;
                    lstCommands.ScrollIntoView(lstCommands.SelectedItem);
                }
                e.Handled = true;
            }
            else if (e.Key == WpfKey.Up)
            {
                if (lstCommands.SelectedIndex > 0)
                {
                    lstCommands.SelectedIndex--;
                    lstCommands.ScrollIntoView(lstCommands.SelectedItem);
                }
                e.Handled = true;
            }
            else if (e.Key == WpfKey.Enter)
            {
                ExecuteCurrentSelection();
                e.Handled = true;
            }
            else if (e.Key == WpfKey.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void OnWindowKeyDown(object sender, WpfKeyEventArgs e)
        {
            if (e.Key == WpfKey.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void OnWindowDeactivated(object? sender, EventArgs e)
        {
            // Tự động đóng khi người dùng click ra ngoài màn hình Excel
            if (!_isExecuting)
            {
                try
                {
                    Close();
                }
                catch { }
            }
        }

        private void OnCommandSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstCommands.SelectedItem != null)
            {
                lstCommands.ScrollIntoView(lstCommands.SelectedItem);
            }
        }

        private void OnCommandDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ExecuteCurrentSelection();
        }

        private void OnEscClicked(object sender, MouseButtonEventArgs e)
        {
            Close();
        }

        private void ExecuteCurrentSelection()
        {
            if (lstCommands.SelectedItem is PaletteCommandItem selectedItem)
            {
                _isExecuting = true;
                CommandSelected?.Invoke(selectedItem);
                Close();
            }
        }
    }
}
