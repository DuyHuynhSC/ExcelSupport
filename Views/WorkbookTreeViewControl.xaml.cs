using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelSupport.ViewModels;
using WpfContextMenu = System.Windows.Controls.ContextMenu;
using WpfUserControl = System.Windows.Controls.UserControl;

namespace ExcelSupport.Views
{
    public partial class WorkbookTreeViewControl : WpfUserControl
    {
        public WorkbookTreeViewControl()
        {
            InitializeComponent();
        }

        protected override void OnPreviewKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            base.OnPreviewKeyDown(e);

            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (Keyboard.FocusedElement is System.Windows.Controls.TextBox tb)
                {
                    if (e.Key == Key.X)
                    {
                        tb.Cut();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.C)
                    {
                        tb.Copy();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.V)
                    {
                        tb.Paste();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.A)
                    {
                        tb.SelectAll();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Z)
                    {
                        if (tb.CanUndo) tb.Undo();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Y)
                    {
                        if (tb.CanRedo) tb.Redo();
                        e.Handled = true;
                    }
                }
            }
        }

        private void OnWorksheetListBoxPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject dep)
            {
                var item = FindAncestor<ListBoxItem>(dep);
                if (item != null && item.DataContext is WorksheetNodeViewModel ws)
                {
                    item.IsSelected = true;
                    if (DataContext is TaskPaneViewModel vm)
                    {
                        vm.SelectedWorksheet = ws;
                        vm.ActivateWorksheetCommand.Execute(ws);
                    }
                }
            }
        }

        private void OnWorkbookListBoxPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject dep)
            {
                var item = FindAncestor<ListBoxItem>(dep);
                if (item != null && item.DataContext is WorkbookNodeViewModel wb)
                {
                    item.IsSelected = true;
                    if (DataContext is TaskPaneViewModel vm)
                    {
                        vm.SelectedWorkbook = wb;
                        vm.ActivateWorkbookCommand.Execute(wb);
                    }
                }
            }
        }

        private void OnListBoxPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject dep)
            {
                var item = FindAncestor<ListBoxItem>(dep);
                if (item != null)
                {
                    item.IsSelected = true;
                }
            }
        }

        private void OnSelectSheetsTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is TaskPaneViewModel vm)
            {
                vm.SelectedTabIndex = 0;
            }
        }

        private void OnSelectAiAssistantTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is TaskPaneViewModel vm)
            {
                vm.SelectedTabIndex = 1;
            }
        }

        private void OnSelectAiSettingsTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is TaskPaneViewModel vm)
            {
                vm.SelectedTabIndex = 2;
            }
        }

        private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                {
                    return match;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }
    }
}
