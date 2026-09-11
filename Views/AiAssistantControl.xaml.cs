using System.Windows;
using ExcelSupport.ViewModels;
using WpfUserControl = System.Windows.Controls.UserControl;

namespace ExcelSupport.Views
{
    public partial class AiAssistantControl : WpfUserControl
    {
        public AiAssistantControl()
        {
            InitializeComponent();
            DataContextChanged += OnDataContextChanged;
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.OldValue is AiAssistantViewModel oldVm)
            {
                oldVm.SheetChatMessages.CollectionChanged -= OnSheetChatMessagesCollectionChanged;
            }

            if (e.NewValue is AiAssistantViewModel newVm)
            {
                newVm.SheetChatMessages.CollectionChanged += OnSheetChatMessagesCollectionChanged;
            }
        }

        private void OnSheetChatMessagesCollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            Dispatcher.InvokeAsync(() =>
            {
                SheetChatScrollViewer?.ScrollToBottom();
            });
        }

        private void OnSheetChatInputPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == 0)
            {
                if (DataContext is AiAssistantViewModel vm && vm.SendSheetChatCommand.CanExecute(null))
                {
                    vm.SendSheetChatCommand.Execute(null);
                    e.Handled = true;
                }
            }
        }

        private void OnSelectFormulaTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.SelectedSubTab = 0;
            }
        }

        private void OnSelectDebugTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.SelectedSubTab = 1;
            }
        }

        private void OnPillSumifsClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.FormulaPrompt = "Tính tổng cột D nếu cột A là 'Đã duyệt' và cột B lớn hơn 100";
            }
        }

        private void OnPillXlookupClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.FormulaPrompt = "Tìm kiếm giá trị ở cột A trong bảng tham chiếu C:E và trả về giá trị ở cột E (nếu không tìm thấy thì trả về rỗng)";
            }
        }

        private void OnPillSplitTextClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.FormulaPrompt = "Tách lấy phần Họ (chữ đầu tiên) và Tên (chữ cuối cùng) từ chuỗi họ tên đầy đủ ở ô A2";
            }
        }
    }
}
