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
