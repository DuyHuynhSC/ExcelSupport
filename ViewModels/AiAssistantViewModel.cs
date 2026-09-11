using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelSupport.Models;
using ExcelSupport.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelSupport.ViewModels
{
    public class AiAssistantViewModel : ViewModelBase
    {
        private int _selectedSubTab = 0; // 0: Sinh công thức, 1: Gỡ lỗi & Hỏi đáp
        private bool _isBusy;
        private string _statusMessage = string.Empty;

        // --- Chat With Sheet Properties ---
        public ObservableCollection<ChatSheetMessageItem> SheetChatMessages { get; } = new ObservableCollection<ChatSheetMessageItem>();

        private string _sheetChatInput = string.Empty;
        public string SheetChatInput
        {
            get => _sheetChatInput;
            set => SetProperty(ref _sheetChatInput, value);
        }

        private bool _isSelectionOnly;
        public bool IsSelectionOnly
        {
            get => _isSelectionOnly;
            set => SetProperty(ref _isSelectionOnly, value);
        }

        public bool HasSheetChatMessages => SheetChatMessages.Count > 0;

        // --- Formula Generator Properties ---
        private string _formulaPrompt = string.Empty;
        private string _formulaResponse = string.Empty;
        private string _extractedFormula = string.Empty;
        private string _formulaExplanation = string.Empty;

        // --- Cell Inspector & Q&A Properties ---
        private AddInEvents.ActiveCellInfo? _activeCell;
        private string _cellInspectorSummary = string.Empty;
        private string _chatPrompt = string.Empty;
        private string _chatResponse = string.Empty;
        private string _fixedFormula = string.Empty;

        public string FixedFormula
        {
            get => _fixedFormula;
            private set
            {
                if (SetProperty(ref _fixedFormula, value))
                {
                    OnPropertyChanged(nameof(HasFixedFormula));
                }
            }
        }

        public bool HasFixedFormula => !string.IsNullOrWhiteSpace(FixedFormula);

        public int SelectedSubTab
        {
            get => _selectedSubTab;
            set => SetProperty(ref _selectedSubTab, value);
        }

        public bool IsBusy
        {
            get => _isBusy;
            private set => SetProperty(ref _isBusy, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            private set => SetProperty(ref _statusMessage, value);
        }

        public string FormulaPrompt
        {
            get => _formulaPrompt;
            set => SetProperty(ref _formulaPrompt, value);
        }

        public string FormulaResponse
        {
            get => _formulaResponse;
            private set
            {
                if (SetProperty(ref _formulaResponse, value))
                {
                    OnPropertyChanged(nameof(HasFormulaResponse));
                }
            }
        }

        public bool HasFormulaResponse => !string.IsNullOrWhiteSpace(FormulaResponse);

        public string ExtractedFormula
        {
            get => _extractedFormula;
            private set
            {
                if (SetProperty(ref _extractedFormula, value))
                {
                    OnPropertyChanged(nameof(HasExtractedFormula));
                }
            }
        }

        public bool HasExtractedFormula => !string.IsNullOrWhiteSpace(ExtractedFormula);

        public string FormulaExplanation
        {
            get => _formulaExplanation;
            private set => SetProperty(ref _formulaExplanation, value);
        }

        public AddInEvents.ActiveCellInfo? ActiveCell
        {
            get => _activeCell;
            private set
            {
                if (SetProperty(ref _activeCell, value))
                {
                    OnPropertyChanged(nameof(HasActiveCell));
                }
            }
        }

        public bool HasActiveCell => ActiveCell != null;

        public string CellInspectorSummary
        {
            get => _cellInspectorSummary;
            private set
            {
                if (SetProperty(ref _cellInspectorSummary, value))
                {
                    OnPropertyChanged(nameof(HasCellInspectorSummary));
                }
            }
        }

        public bool HasCellInspectorSummary => !string.IsNullOrWhiteSpace(CellInspectorSummary);

        public string ChatPrompt
        {
            get => _chatPrompt;
            set => SetProperty(ref _chatPrompt, value);
        }

        public string ChatResponse
        {
            get => _chatResponse;
            private set
            {
                if (SetProperty(ref _chatResponse, value))
                {
                    OnPropertyChanged(nameof(HasChatResponse));
                }
            }
        }

        public bool HasChatResponse => !string.IsNullOrWhiteSpace(ChatResponse);

        // --- Commands ---
        public ICommand GenerateFormulaCommand { get; }
        public ICommand InsertFormulaToExcelCommand { get; }
        public ICommand CopyFormulaCommand { get; }
        public ICommand ReadActiveCellCommand { get; }
        public ICommand DebugActiveCellCommand { get; }
        public ICommand ApplyFixFormulaCommand { get; }
        public ICommand SendChatCommand { get; }
        public ICommand ClearFormulaCommand { get; }
        public ICommand ClearChatCommand { get; }
        public ICommand CopyChatResponseCommand { get; }

        // --- Chat With Sheet Commands ---
        public ICommand SendSheetChatCommand { get; }
        public ICommand ClearSheetChatCommand { get; }
        public ICommand NavigateToCellCommand { get; }
        public ICommand ApplySuggestedFormulaCommand { get; }
        public ICommand QuickPromptCommand { get; }

        public AiAssistantViewModel()
        {
            GenerateFormulaCommand = new RelayCommand(async _ => await ExecuteGenerateFormulaAsync(), _ => !IsBusy && !string.IsNullOrWhiteSpace(FormulaPrompt));
            InsertFormulaToExcelCommand = new RelayCommand(_ => ExecuteInsertFormula());
            CopyFormulaCommand = new RelayCommand(_ => ExecuteCopyFormula());
            ReadActiveCellCommand = new RelayCommand(_ => ExecuteReadActiveCell());
            DebugActiveCellCommand = new RelayCommand(async _ => await ExecuteDebugActiveCellAsync(), _ => !IsBusy);
            ApplyFixFormulaCommand = new RelayCommand(_ => ExecuteApplyFixFormula());
            SendChatCommand = new RelayCommand(async _ => await ExecuteSendChatAsync(), _ => !IsBusy && !string.IsNullOrWhiteSpace(ChatPrompt));
            ClearFormulaCommand = new RelayCommand(_ => ExecuteClearFormula());
            ClearChatCommand = new RelayCommand(_ => ExecuteClearChat());
            CopyChatResponseCommand = new RelayCommand(_ => ExecuteCopyChatResponse());

            SendSheetChatCommand = new RelayCommand(async _ => await ExecuteSendSheetChatAsync(), _ => !IsBusy && !string.IsNullOrWhiteSpace(SheetChatInput));
            ClearSheetChatCommand = new RelayCommand(_ => ExecuteClearSheetChat());
            NavigateToCellCommand = new RelayCommand(param => ExecuteNavigateToCell(param as string));
            ApplySuggestedFormulaCommand = new RelayCommand(param => ExecuteApplySuggestedFormula(param as string));
            QuickPromptCommand = new RelayCommand(async param => await ExecuteQuickPromptAsync(param as string));

            InitWelcomeMessage();
        }

        #region Formula Generator Logic

        private async Task ExecuteGenerateFormulaAsync()
        {
            if (string.IsNullOrWhiteSpace(FormulaPrompt)) return;

            IsBusy = true;
            StatusMessage = "AI đang suy nghĩ và viết công thức... ⏳";
            FormulaResponse = string.Empty;
            ExtractedFormula = string.Empty;
            FormulaExplanation = string.Empty;

            try
            {
                var config = AiConfigManager.Current;
                string sheetContext = AddInEvents.Instance?.GetActiveSheetContextSummary() ?? string.Empty;
                string contextPart = !string.IsNullOrEmpty(sheetContext) ? $"\n\n[NGỮ CẢNH BẢNG TÍNH HIỆN TẠI]:\n{sheetContext}" : "";

                string systemPrompt = "Bạn là chuyên gia Excel hàng đầu thế giới. Người dùng sẽ yêu cầu bạn tạo một công thức Excel từ mô tả tiếng Việt.\n" +
                                      "Yêu cầu:\n" +
                                      "1. Đưa ra công thức Excel CHUẨN XÁC NHẤT bắt đầu bằng dấu '=' và đặt bên trong khối code: ```excel\n=CÔNG_THỨC\n```\n" +
                                      "2. Sử dụng chính xác chữ cái cột (A, B, C, ...) dựa trên cấu trúc bảng tính được cung cấp.\n" +
                                      "3. Ngay bên dưới khối code, hãy giải thích ngắn gọn, dễ hiểu bằng tiếng Việt về cách hoạt động của từng tham số trong công thức.\n" +
                                      "4. Ưu tiên các hàm hiện đại như XLOOKUP, SUMIFS, FILTER, UNIQUE, TEXTJOIN nếu phù hợp.";

                string fullPrompt = FormulaPrompt + contextPart;
                string response = await Task.Run(() => OpenAiClientService.SendChatAsync(config, fullPrompt, systemPrompt));

                FormulaResponse = response;
                ExtractedFormula = ExtractFormula(response);
                FormulaExplanation = ExtractExplanation(response);
            }
            catch (Exception ex)
            {
                FormulaResponse = $"❌ Lỗi khi sinh công thức: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                StatusMessage = string.Empty;
            }
        }

        private static string ExtractFormula(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return string.Empty;

            // 1. Tìm trong ```excel ... ``` hoặc ``` ... ```
            var codeMatch = Regex.Match(response, @"```(?:excel|plaintext)?\s*(=[^\n`]+)\s*```", RegexOptions.IgnoreCase);
            if (codeMatch.Success)
            {
                return codeMatch.Groups[1].Value.Trim();
            }

            // 2. Tìm dòng có dạng 🛠️ Công thức sửa: =...
            var fixMatch = Regex.Match(response, @"(?:Công thức sửa|Công thức đề xuất|Công thức)\s*:\s*(=[A-Z_]+(?:\.[A-Z_]+)?\(.*?\)|=[^\r\n]+)", RegexOptions.IgnoreCase);
            if (fixMatch.Success)
            {
                return fixMatch.Groups[1].Value.Trim();
            }

            // 3. Tìm dòng bất kỳ bắt đầu bằng '=' có chứa hàm Excel
            var lineMatch = Regex.Match(response, @"(?m)^\s*(=[A-Z_]+(?:\.[A-Z_]+)?\(.*?\))\s*$", RegexOptions.Multiline);
            if (lineMatch.Success)
            {
                return lineMatch.Groups[1].Value.Trim();
            }

            // 4. Tìm bất kỳ biểu thức =FUNCTION(...)
            var funcMatch = Regex.Match(response, @"(=[A-Z_]{2,}\([^\r\n]*\))");
            if (funcMatch.Success)
            {
                return funcMatch.Groups[1].Value.Trim();
            }

            return string.Empty;
        }

        private static string ExtractExplanation(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return string.Empty;
            // Bỏ khối code và giữ phần giải thích
            string cleaned = Regex.Replace(response, @"```(?:excel|plaintext)?\s*(=[^\n`]+)\s*```", "", RegexOptions.IgnoreCase).Trim();
            return cleaned;
        }

        private void ExecuteInsertFormula()
        {
            if (string.IsNullOrWhiteSpace(ExtractedFormula)) return;

            var addIn = AddInEvents.Instance;
            if (addIn != null)
            {
                bool ok = addIn.InsertFormulaToActiveCell(ExtractedFormula);
                if (ok)
                {
                    StatusMessage = "⚡ Đã chèn công thức vào ô đang chọn!";
                }
            }
        }

        private void ExecuteApplyFixFormula()
        {
            if (string.IsNullOrWhiteSpace(FixedFormula)) return;

            var addIn = AddInEvents.Instance;
            if (addIn != null)
            {
                bool ok = addIn.InsertFormulaToActiveCell(FixedFormula);
                if (ok)
                {
                    StatusMessage = "🛠️ Đã áp dụng công thức sửa lỗi vào ô Excel thành công!";
                    ExecuteReadActiveCell(); // Cập nhật lại trạng thái ô
                }
            }
        }

        private void ExecuteCopyFormula()
        {
            if (!string.IsNullOrWhiteSpace(ExtractedFormula))
            {
                try
                {
                    System.Windows.Clipboard.SetText(ExtractedFormula);
                    StatusMessage = "📋 Đã sao chép công thức vào Clipboard!";
                }
                catch { }
            }
        }

        private void ExecuteClearFormula()
        {
            FormulaPrompt = string.Empty;
            FormulaResponse = string.Empty;
            ExtractedFormula = string.Empty;
            FormulaExplanation = string.Empty;
            StatusMessage = string.Empty;
        }

        #endregion

        #region Cell Inspector & Debugger Logic

        private void ExecuteReadActiveCell()
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            var info = addIn.GetActiveCellInfo();
            ActiveCell = info;

            if (info != null)
            {
                string formulaText = !string.IsNullOrEmpty(info.Formula) ? info.Formula : "(Không có công thức)";
                string valText = !string.IsNullOrEmpty(info.Value) ? info.Value : "(Trống)";
                string errNotice = info.HasError ? $"\n⚠️ Ô đang gặp mã lỗi: {info.ErrorText}" : "";

                CellInspectorSummary = $"📍 Ô: [{info.SheetName}!{info.CellAddress}]\n" +
                                       $"🔹 Giá trị: {valText}\n" +
                                       $"🔹 Công thức: {formulaText}{errNotice}";
            }
            else
            {
                CellInspectorSummary = "⚠️ Không thể đọc ô đang chọn. Vui lòng nhấp vào một ô trên bảng tính Excel.";
            }
        }

        private async Task ExecuteDebugActiveCellAsync()
        {
            if (ActiveCell == null)
            {
                ExecuteReadActiveCell();
                if (ActiveCell == null) return;
            }

            IsBusy = true;
            StatusMessage = "AI đang kiểm tra và phân tích ô tính... ⏳";
            ChatResponse = string.Empty;
            FixedFormula = string.Empty;

            try
            {
                var config = AiConfigManager.Current;
                string sheetContext = AddInEvents.Instance?.GetActiveSheetContextSummary() ?? string.Empty;
                string contextPart = !string.IsNullOrEmpty(sheetContext) ? $"\n\n[NGỮ CẢNH BẢNG TÍNH]:\n{sheetContext}" : "";

                string prompt = $"Phân tích và gỡ lỗi ô Excel sau:\n" +
                               $"- Sheet: {ActiveCell.SheetName}\n" +
                               $"- Tọa độ ô: {ActiveCell.CellAddress}\n" +
                               $"- Công thức: {ActiveCell.Formula}\n" +
                               $"- Giá trị hiển thị: {ActiveCell.Value}\n" +
                               $"- Lỗi: {(ActiveCell.HasError ? ActiveCell.ErrorText : "Không có")}{contextPart}\n\n" +
                               $"Hãy giải thích nguyên nhân và đưa ra công thức sửa lại hoàn chỉnh.";

                string systemPrompt = 
                    "Bạn là chuyên gia kiểm tra và gỡ lỗi công thức Excel trên thanh công cụ TaskPane.\n" +
                    "YÊU CẦU ĐỊNH DẠNG (BẮT BUỘC):\n" +
                    "1. Trả lời CỰC KỲ NGẮN GỌN (dưới 8 dòng), đi thẳng vào vấn đề.\n" +
                    "2. KHÔNG dùng cú pháp markdown thô (như ###, ```, **).\n" +
                    "3. Trình bày trực quan với biểu tượng:\n" +
                    "   🔍 Phân tích: (ngắn gọn 1-2 câu)\n" +
                    "   🛠️ Công thức sửa: =CÔNG_THỨC (công thức Excel chuẩn để người dùng áp dụng ngay)\n" +
                    "   💡 Giải thích: (1 câu lý do ngắn gọn)";

                string reply = await Task.Run(() => OpenAiClientService.SendChatAsync(config, prompt, systemPrompt));
                ChatResponse = FormatAiResponseForUi(reply);
                FixedFormula = ExtractFormula(reply);
                SelectedSubTab = 2; // Chuyển sang xem phản hồi ở tab Gỡ lỗi
            }
            catch (Exception ex)
            {
                ChatResponse = $"❌ Lỗi phân tích: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                StatusMessage = string.Empty;
            }
        }

        private async Task ExecuteSendChatAsync()
        {
            if (string.IsNullOrWhiteSpace(ChatPrompt)) return;

            IsBusy = true;
            StatusMessage = "AI đang suy nghĩ và trả lời... ⏳";
            ChatResponse = string.Empty;

            try
            {
                var config = AiConfigManager.Current;
                string systemPrompt = 
                    "Bạn là trợ lý AI thông minh chuyên về Excel trên thanh công cụ hỗ trợ người dùng.\n" +
                    "YÊU CẦU ĐỊNH DẠNG & TRÌNH BÀY (BẮT BUỘC):\n" +
                    "1. Trả lời CỰC KỲ NGẮN GỌN, súc tích (khoảng 3 - 6 dòng), đi thẳng vào câu trả lời, không chào hỏi dài dòng.\n" +
                    "2. KHÔNG sử dụng ký tự Markdown thô (như ###, ```, **, *).\n" +
                    "3. Trình bày trực quan, rõ ràng bằng các biểu tượng:\n" +
                    "   📌 Ý nghĩa: (1-2 câu giải thích ngắn gọn)\n" +
                    "   🔢 Kết quả: (nêu kết quả với dữ liệu hiện tại nếu có)\n" +
                    "   💡 Lưu ý: (mẹo hoặc lưu ý quan trọng nếu có)\n" +
                    "4. Dùng tiếng Việt chuẩn, tự nhiên, dễ hiểu.";

                // Tự động đính kèm ngữ cảnh ô đang chọn nếu có
                var addIn = AddInEvents.Instance;
                var currentCell = addIn?.GetActiveCellInfo() ?? ActiveCell;
                string contextInfo = "";
                if (currentCell != null && (!string.IsNullOrEmpty(currentCell.Formula) || !string.IsNullOrEmpty(currentCell.Value)))
                {
                    contextInfo = $"\n\n[NGỮ CẢNH Ô ĐANG CHỌN TRÊN EXCEL]:\n" +
                                  $"- Tọa độ ô: [{currentCell.SheetName}!{currentCell.CellAddress}]\n" +
                                  $"- Công thức trong ô: {(!string.IsNullOrEmpty(currentCell.Formula) ? currentCell.Formula : "(Không có công thức)")}\n" +
                                  $"- Giá trị hiển thị: {(!string.IsNullOrEmpty(currentCell.Value) ? currentCell.Value : "(Trống)")}\n" +
                                  $"{(currentCell.HasError ? $"- Trạng thái lỗi: {currentCell.ErrorText}\n" : "")}";
                }

                string fullPrompt = ChatPrompt + contextInfo;
                string reply = await Task.Run(() => OpenAiClientService.SendChatAsync(config, fullPrompt, systemPrompt));
                ChatResponse = FormatAiResponseForUi(reply);
            }
            catch (Exception ex)
            {
                ChatResponse = $"❌ Lỗi: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                StatusMessage = string.Empty;
            }
        }

        private void ExecuteCopyChatResponse()
        {
            if (!string.IsNullOrWhiteSpace(ChatResponse))
            {
                try
                {
                    System.Windows.Clipboard.SetText(ChatResponse);
                    StatusMessage = "📋 Đã sao chép phản hồi vào Clipboard!";
                }
                catch { }
            }
        }

        private void ExecuteClearChat()
        {
            ChatPrompt = string.Empty;
            ChatResponse = string.Empty;
            StatusMessage = string.Empty;
        }

        public static string FormatAiResponseForUi(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            string s = text!;

            // 1. Chuyển đổi khối code markdown
            s = Regex.Replace(s, @"```(?:excel|plaintext)?\s*\n?(=[^\r\n`]+)\s*\n?```", "👉 $1", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"```[a-zA-Z]*\s*\n?([\s\S]*?)\n?```", "$1");

            // 2. Chuyển đổi các tiêu đề Markdown thành icon trực quan
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Ý nghĩa|Mục đích)[:\s]*", "📌 Ý nghĩa:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Với giá trị hiện tại|Giá trị hiện tại|Cách tính)[:\s]*", "🔢 Giá trị & Cách tính:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Lưu ý|Chú ý)[:\s]*", "💡 Lưu ý:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Gợi ý|Đề xuất)[:\s]*", "⚡ Gợi ý:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Nguyên nhân|Lỗi)[:\s]*", "🔍 Nguyên nhân:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(?:Cách khắc phục|Sửa lỗi)[:\s]*", "🛠️ Cách khắc phục:\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?m)^\s*#{1,4}\s*(.*?)\s*$", "🔹 $1:", RegexOptions.Multiline);

            // 3. Chuyển đổi gạch đầu dòng markdown thành bullet đẹp
            s = Regex.Replace(s, @"(?m)^\s*[\-\*]\s+", "  • ", RegexOptions.Multiline);

            // 4. Loại bỏ dấu in đậm và code thô
            s = Regex.Replace(s, @"\*\*([^\*]+)\*\*", "$1");
            s = Regex.Replace(s, @"\*([^\*]+)\*", "$1");
            s = Regex.Replace(s, @"`([^`]+)`", "$1");

            // 5. Rút gọn dòng trống liên tiếp
            s = Regex.Replace(s, @"(\r?\n){3,}", "\n\n");

            return s.Trim();
        }

        #endregion

        #region Chat With Sheet Logic

        private void InitWelcomeMessage()
        {
            SheetChatMessages.Clear();
            SheetChatMessages.Add(new ChatSheetMessageItem
            {
                IsUser = false,
                Content = "👋 Xin chào! Tôi là Trợ lý AI Bảng tính (Chat With Sheet).\n\n" +
                          "Bạn có thể hỏi bất kỳ điều gì về bảng tính hiện tại, ví dụ:\n" +
                          "• Tóm tắt cấu trúc bảng & các cột dữ liệu\n" +
                          "• Tìm kiếm giá trị lớn nhất, nhỏ nhất, bất thường\n" +
                          "• Viết công thức phân tích hoặc tính toán theo điều kiện\n\n" +
                          "💡 Mẹo: Bạn có thể click vào các tọa độ ô được AI nhắc tới để Excel tự động nhảy đến ô đó!"
            });
            OnPropertyChanged(nameof(HasSheetChatMessages));
        }

        private void ExecuteClearSheetChat()
        {
            InitWelcomeMessage();
            SheetChatInput = string.Empty;
            StatusMessage = string.Empty;
        }

        private void ExecuteNavigateToCell(string? cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress)) return;
            var addIn = AddInEvents.Instance;
            if (addIn != null)
            {
                addIn.NavigateToCell(cellAddress!);
            }
        }

        private void ExecuteApplySuggestedFormula(string? formula)
        {
            string targetFormula = !string.IsNullOrWhiteSpace(formula) ? formula! : FixedFormula;
            if (string.IsNullOrWhiteSpace(targetFormula)) return;

            var addIn = AddInEvents.Instance;
            if (addIn != null)
            {
                bool ok = addIn.InsertFormulaToActiveCell(targetFormula);
                if (ok)
                {
                    StatusMessage = "🛠️ Đã chèn công thức đề xuất vào ô đang chọn!";
                }
            }
        }

        private async Task ExecuteQuickPromptAsync(string? prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return;
            SheetChatInput = prompt!;
            await ExecuteSendSheetChatAsync();
        }

        private async Task ExecuteSendSheetChatAsync()
        {
            string userQuestion = SheetChatInput?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(userQuestion)) return;

            // 1. Thêm tin nhắn của người dùng vào giao diện ngay lập tức
            var userItem = new ChatSheetMessageItem
            {
                IsUser = true,
                Content = userQuestion
            };
            SheetChatMessages.Add(userItem);
            SheetChatInput = string.Empty;
            OnPropertyChanged(nameof(HasSheetChatMessages));

            IsBusy = true;
            StatusMessage = "AI đang đọc dữ liệu Sheet và suy nghĩ... ⏳";

            try
            {
                var config = AiConfigManager.Current;
                var addIn = AddInEvents.Instance;
                string sheetData = addIn?.GetSheetChatContext(IsSelectionOnly) ?? string.Empty;

                string systemPrompt = 
                    "Bạn là Trợ lý AI Bảng tính (Chat With Sheet) cao cấp, chuyên gia phân tích dữ liệu Excel.\n" +
                    "Bạn giao tiếp bằng tiếng Việt tự nhiên, ngắn gọn, lịch sự, chuyên nghiệp và đi thẳng vào trọng tâm.\n\n" +
                    "QUY TẮC QUAN TRỌNG (BẮT BUỘC TUÂN THỦ):\n" +
                    "1. Dữ liệu bảng tính hiện tại được cung cấp dưới dạng Markdown. Hãy căn cứ CHÍNH XÁC vào dữ liệu này để trả lời.\n" +
                    "2. Khi nhắc tới bất kỳ ô hoặc vùng ô cụ thể nào, hãy LUÔN đặt trong ngoặc vuông dạng [CộtDòng] (ví dụ: [A1], [B5], [C2:D10], [DoanhThu!C5]) để hệ thống tạo liên kết click nhảy ô cho người dùng.\n" +
                    "3. Nếu người dùng hỏi cách tính toán hoặc đề xuất công thức Excel, hãy cung cấp công thức chuẩn xác bắt đầu bằng dấu '=' trong khối code: ```excel\n=CÔNG_THỨC\n```\n" +
                    "4. Trình bày câu trả lời rõ ràng, dùng bullet point (•), biểu tượng (📌, 🔢, 💡, ⚠️) và in đậm tiêu đề quan trọng.";

                // Chuẩn bị danh sách lịch sử hội thoại (lấy tối đa 6 lượt tin nhắn gần nhất)
                var historyList = new List<(string role, string content)>();

                // Đưa context dữ liệu sheet vào câu hỏi hiện tại hoặc lượt đầu tiên
                string promptWithContext = $"[DỮ LIỆU BẢNG TÍNH HIỆN TẠI]:\n{sheetData}\n\n[CÂU HỎI CỦA NGƯỜI DÙNG]:\n{userQuestion}";

                int historyStartIndex = Math.Max(0, SheetChatMessages.Count - 7);
                for (int i = historyStartIndex; i < SheetChatMessages.Count - 1; i++)
                {
                    var msg = SheetChatMessages[i];
                    if (msg.IsUser)
                    {
                        historyList.Add(("user", msg.Content));
                    }
                    else
                    {
                        historyList.Add(("assistant", msg.Content));
                    }
                }

                // Lượt hiện tại kèm ngữ cảnh bảng tính
                historyList.Add(("user", promptWithContext));

                string reply = await Task.Run(() => OpenAiClientService.SendChatMessagesAsync(config, historyList, systemPrompt));

                var assistantItem = new ChatSheetMessageItem
                {
                    IsUser = false,
                    Content = reply.Trim()
                };

                // Trích xuất các ô được nhắc tới
                ExtractReferencedCells(reply, assistantItem);

                // Trích xuất công thức đề xuất nếu có
                string suggestedFormula = ExtractFormula(reply);
                if (!string.IsNullOrWhiteSpace(suggestedFormula))
                {
                    assistantItem.SuggestedFormula = suggestedFormula;
                }

                SheetChatMessages.Add(assistantItem);
                OnPropertyChanged(nameof(HasSheetChatMessages));
            }
            catch (Exception ex)
            {
                SheetChatMessages.Add(new ChatSheetMessageItem
                {
                    IsUser = false,
                    Content = $"❌ Lỗi khi hỏi đáp với Sheet: {ex.Message}"
                });
                OnPropertyChanged(nameof(HasSheetChatMessages));
            }
            finally
            {
                IsBusy = false;
                StatusMessage = string.Empty;
            }
        }

        private static void ExtractReferencedCells(string text, ChatSheetMessageItem item)
        {
            if (string.IsNullOrWhiteSpace(text) || item == null) return;

            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Pattern 1: [A1] hoặc [Sheet1!A1:B10] hoặc ['Tên Sheet'!C5]
            var bracketMatches = Regex.Matches(text, @"\[((?:'[^']+'!|[A-Za-z0-9_]+!)?\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)\]");
            foreach (Match m in bracketMatches)
            {
                string cell = m.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(cell) && added.Add(cell))
                {
                    item.ReferencedCells.Add(cell);
                }
            }

            // Pattern 2: ô A1 hoặc ô B5:C10
            var wordMatches = Regex.Matches(text, @"(?:\bô|\bvùng)\s+([A-Za-z]{1,3}[0-9]+(?::[A-Za-z]{1,3}[0-9]+)?)\b", RegexOptions.IgnoreCase);
            foreach (Match m in wordMatches)
            {
                string cell = m.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(cell) && added.Add(cell))
                {
                    item.ReferencedCells.Add(cell);
                }
            }

            item.NotifyCellsChanged();
        }

        #endregion
    }
}
