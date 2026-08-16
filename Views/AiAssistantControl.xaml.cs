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
        }

        private void OnSelectTranslationTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.SelectedSubTab = 0;
            }
        }

        private void OnSelectFormulaTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.SelectedSubTab = 1;
            }
        }

        private void OnSelectDebugTabClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is AiAssistantViewModel vm)
            {
                vm.SelectedSubTab = 2;
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
