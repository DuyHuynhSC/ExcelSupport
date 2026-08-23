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
                        return;
                    }
                    else if (e.Key == Key.C)
                    {
                        tb.Copy();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.V)
                    {
                        tb.Paste();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.A)
                    {
                        tb.SelectAll();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.Z)
                    {
                        if (tb.CanUndo) tb.Undo();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.Y)
                    {
                        if (tb.CanRedo) tb.Redo();
                        e.Handled = true;
                        return;
                    }
                }
            }

            // Điều hướng toàn cục cho TaskPane khi Tab = 0 (Sheet Navigator): hỗ trợ ↓ / ↑ / Enter mọi lúc mọi nơi
            if (DataContext is TaskPaneViewModel vm && vm.SelectedTabIndex == 0)
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Enter)
                {
                    bool isWorkbookFocus = Keyboard.FocusedElement is System.Windows.Controls.TextBox focusedTb && 
                                           focusedTb.Name == "TxtWorkbookSearch";

                    if (isWorkbookFocus)
                    {
                        HandleWorkbookNavigation(vm, e);
                    }
                    else
                    {
                        HandleWorksheetNavigation(vm, e);
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

        private void OnWorksheetListBoxSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.ListBox lb && lb.SelectedItem is WorksheetNodeViewModel ws)
            {
                if (DataContext is TaskPaneViewModel vm)
                {
                    vm.ActivateWorksheetCommand.Execute(ws);
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

        private void OnWorkbookListBoxSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.ListBox lb && lb.SelectedItem is WorkbookNodeViewModel wb)
            {
                if (DataContext is TaskPaneViewModel vm)
                {
                    vm.ActivateWorkbookCommand.Execute(wb);
                }
            }
        }

        private void OnWorkbookSearchPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (DataContext is TaskPaneViewModel vm)
            {
                HandleWorkbookNavigation(vm, e);
            }
        }

        private void OnSheetSearchPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (DataContext is TaskPaneViewModel vm)
            {
                HandleWorksheetNavigation(vm, e);
            }
        }

        private void HandleWorksheetNavigation(TaskPaneViewModel vm, System.Windows.Input.KeyEventArgs e)
        {
            if (vm.SheetsView == null) return;

            var items = System.Linq.Enumerable.ToList(System.Linq.Enumerable.Cast<WorksheetNodeViewModel>(vm.SheetsView));
            if (items.Count == 0) return;

            int currentIndex = items.FindIndex(s => string.Equals(s.SheetName, vm.SelectedWorksheet?.SheetName, StringComparison.OrdinalIgnoreCase));

            if (e.Key == Key.Down)
            {
                int nextIndex = (currentIndex < 0) ? 0 : Math.Min(items.Count - 1, currentIndex + 1);
                var nextWs = items[nextIndex];
                vm.SelectedWorksheet = nextWs;
                vm.ActivateWorksheetCommand.Execute(nextWs);
                WorksheetsListBox?.ScrollIntoView(nextWs);
                e.Handled = true;
            }
            else if (e.Key == Key.Up)
            {
                int prevIndex = (currentIndex < 0) ? items.Count - 1 : Math.Max(0, currentIndex - 1);
                var prevWs = items[prevIndex];
                vm.SelectedWorksheet = prevWs;
                vm.ActivateWorksheetCommand.Execute(prevWs);
                WorksheetsListBox?.ScrollIntoView(prevWs);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                var targetWs = (currentIndex >= 0) ? items[currentIndex] : items[0];
                vm.SelectedWorksheet = targetWs;
                vm.ActivateWorksheetCommand.Execute(targetWs);
                WorksheetsListBox?.ScrollIntoView(targetWs);
                e.Handled = true;
            }
        }

        private void HandleWorkbookNavigation(TaskPaneViewModel vm, System.Windows.Input.KeyEventArgs e)
        {
            if (vm.WorkbooksView == null) return;

            var items = System.Linq.Enumerable.ToList(System.Linq.Enumerable.Cast<WorkbookNodeViewModel>(vm.WorkbooksView));
            if (items.Count == 0) return;

            int currentIndex = items.FindIndex(w => string.Equals(w.WorkbookName, vm.SelectedWorkbook?.WorkbookName, StringComparison.OrdinalIgnoreCase));

            if (e.Key == Key.Down)
            {
                int nextIndex = (currentIndex < 0) ? 0 : Math.Min(items.Count - 1, currentIndex + 1);
                var nextWb = items[nextIndex];
                vm.SelectedWorkbook = nextWb;
                vm.ActivateWorkbookCommand.Execute(nextWb);
                WorkbooksListBox?.ScrollIntoView(nextWb);
                e.Handled = true;
            }
            else if (e.Key == Key.Up)
            {
                int prevIndex = (currentIndex < 0) ? items.Count - 1 : Math.Max(0, currentIndex - 1);
                var prevWb = items[prevIndex];
                vm.SelectedWorkbook = prevWb;
                vm.ActivateWorkbookCommand.Execute(prevWb);
                WorkbooksListBox?.ScrollIntoView(prevWb);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                var targetWb = (currentIndex >= 0) ? items[currentIndex] : items[0];
                vm.SelectedWorkbook = targetWb;
                vm.ActivateWorkbookCommand.Execute(targetWb);
                WorkbooksListBox?.ScrollIntoView(targetWb);
                e.Handled = true;
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
