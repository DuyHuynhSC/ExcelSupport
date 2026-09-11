# Thiết Kế Chi Tiết: Chat With Sheet (Task Pane)

Tài liệu đặc tả thiết kế kỹ thuật (Design Specification) cho tính năng **Chat With Sheet (Hỏi đáp trực tiếp với Sheet trong Task Pane)** thuộc phân hệ **Next-Gen AI Copilot** (mục 5.3 trong [FEATURE_ROADMAP.md](../../FEATURE_ROADMAP.md)).

---

## 1. Mục Tiêu & Vấn Đề Giải Quyết
* **Vấn đề:** Người dùng Excel thường mất thời gian đọc hiểu các bảng dữ liệu lớn, tính toán nhẩm, tìm kiếm dữ liệu bất thường hoặc viết các công thức tra cứu phức tạp. Khung hỏi đáp hiện tại trong Task Pane chỉ là ô hỏi đáp đơn lẻ một lượt, không có ngữ cảnh hội thoại liên tục, không tự động bóc tách cấu trúc bảng tính và không có liên kết tương tác ngược lại bảng tính Excel.
* **Mục tiêu:** 
  - Cho phép người dùng trò chuyện tương tác (Multi-turn Chat) bằng tiếng Việt tự nhiên với dữ liệu bảng tính đang mở ngay trên Task Pane.
  - Tự động trích xuất thông minh ngữ cảnh dữ liệu từ Sheet (Vùng chọn `Selection` hoặc cấu trúc Header + dữ liệu mẫu của `UsedRange`).
  - Hỗ trợ **nhảy trực tiếp tới ô dữ liệu (Cell Jump Links)**: Khi AI nhắc tới tọa độ ô/vùng (ví dụ `[A1]`, `[B2:D10]`, `[DoanhThu!C5]`), hiển thị dưới dạng badge liên kết; bấm vào là Excel tự động cuộn đến và kích hoạt/chọn ô đó.
  - Hỗ trợ **áp dụng nhanh công thức đề xuất**: 1-Click chèn công thức Excel do AI viết thẳng vào ô đang chọn.
  - Hỗ trợ **Gợi ý câu hỏi nhanh (Quick Prompt Pills)** giúp người dùng kích hoạt phân tích ngay chỉ với 1 click.

---

## 2. Kiến Trúc Tổng Thể & Các Thành Phần

```mermaid
graph TD
    User([Người Dùng]) -->|Hỏi đáp / Click Pill| UI[Chat With Sheet UI - WPF]
    UI -->|Gửi câu hỏi + Lịch sử| VM[AiAssistantViewModel]
    VM -->|Thu thập dữ liệu bảng tính| Extractor[AddInEvents Sheet Context Extractor]
    Extractor -->|Dữ liệu bảng Markdown/JSON| VM
    VM -->|Gửi Multi-turn Messages| Client[OpenAiClientService]
    Client -->|API /chat/completions| LLM[AI Model LLM]
    LLM -->|Phản hồi văn bản + Tọa độ ô [A1]| Client
    Client -->|Kết quả chuỗi| VM
    VM -->|Phân tích Tọa độ & Công thức| Parser[Regex Cell & Formula Parser]
    Parser -->|Tạo Bubble + Badges [A1]| UI
    UI -->|Click badge [A1]| Navigator[AddInEvents Cell Navigation]
    Navigator -->|Excel COM Goto / Select| ExcelSheet[Bảng Tính Excel]
```

---

## 3. Chi Tiết Các Thành Phần Kỹ Thuật

### 3.1. Dữ Liệu & Models (`ExcelSupport.Models`)
Tạo model `ChatSheetMessageItem`:
```csharp
public class ChatSheetMessageItem : ViewModelBase
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public bool IsUser { get; set; }
    public string Content { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; } = DateTime.Now;
    
    // Danh sách các ô/vùng được AI nhắc tới trong câu trả lời (ví dụ: "A1", "B2:D10")
    public ObservableCollection<string> ReferencedCells { get; } = new ObservableCollection<string>();
    public bool HasReferencedCells => ReferencedCells.Count > 0;

    // Công thức Excel được trích xuất từ câu trả lời (nếu có)
    public string? SuggestedFormula { get; set; }
    public bool HasSuggestedFormula => !string.IsNullOrEmpty(SuggestedFormula);
}
```

### 3.2. Thu Thập Ngữ Cảnh Bảng Tính (`AddInEvents.AiHelpers.cs`)
Bổ sung phương thức `GetSheetChatContext(bool selectionOnly, int maxRows = 100, int maxCols = 25)`:
1. **Kiểm tra vùng chọn (`Selection`):**
   - Nếu `selectionOnly == true` hoặc `Selection.CountLarge > 1`:
     - Giới hạn tối đa `maxRows` dòng x `maxCols` cột để tối ưu token.
     - Đọc dữ liệu nhanh bằng mảng 2 chiều 1 COM call (`Value2 as object[,]`).
     - Định dạng thành bảng Markdown rõ ràng kèm tọa độ hàng/cột tương ứng.
2. **Nếu không chọn vùng (chỉ trỏ 1 ô hoặc `selectionOnly == false`):**
   - Đọc thông tin Workbook, Sheet hiện tại, tọa độ ô ActiveCell, giá trị và công thức ô đó.
   - Quét `UsedRange`: lấy dòng tiêu đề (Headers) và tối đa `maxRows` dòng dữ liệu đại diện.
   - Đính kèm tổng số dòng/cột của Sheet để AI có bức tranh tổng thể.

### 3.3. Tương Tác Nhảy Ô & Chèn Công Thức (`AddInEvents.AiHelpers.cs`)
* `NavigateToCell(string cellAddressOrRange)`:
  - Phân tích cú pháp: hỗ trợ cả địa chỉ đơn `A1`, vùng `B2:D10`, hoặc có tên sheet `'Tên Sheet'!A1`.
  - Sử dụng `ExcelAsyncUtil.QueueAsMacro` để đảm bảo luồng Excel an toàn:
    ```csharp
    dynamic app = _excelApp;
    dynamic targetRange = ws.Range[cleanAddress];
    app.Goto(targetRange, true); // Cuộn và kích hoạt ô ngay lập tức
    ```
* `InsertFormulaToActiveCell(string formula)`:
  - Đã có sẵn trong `AddInEvents`, tái sử dụng khi người dùng bấm nút chèn công thức.

### 3.4. Dịch Vụ AI Client (`OpenAiClientService.cs`)
Bổ sung hỗ trợ hội thoại nhiều lượt:
```csharp
public static async Task<string> SendChatMessagesAsync(
    AiConfig config, 
    IEnumerable<(string role, string content)> messages, 
    string? systemPrompt = null)
```
* Tự động duy trì ngữ cảnh tối đa 6 lượt trao đổi gần nhất (3 user, 3 assistant) để tiết kiệm token và đảm bảo tốc độ phản hồi nhanh.

### 3.5. Logic ViewModel (`AiAssistantViewModel.cs`)
* **Thuộc tính mới:**
  - `ObservableCollection<ChatSheetMessageItem> SheetChatMessages`: Danh sách bong bóng hội thoại.
  - `bool IsSelectionOnly`: Toggle chế độ "Chỉ lấy vùng đang chọn" vs "Cả bảng tính".
  - `string SheetChatInput`: Nội dung câu hỏi người dùng đang gõ.
  - `ICommand SendSheetChatCommand`: Gửi câu hỏi kèm dữ liệu bảng tính.
  - `ICommand ClearSheetChatCommand`: Làm mới cuộc trò chuyện.
  - `ICommand NavigateToCellCommand`: Kích hoạt nhảy tới ô trên Excel.
  - `ICommand ApplySuggestedFormulaCommand`: Chèn công thức AI đề xuất vào ô active.
  - `ICommand QuickPromptCommand`: Chạy nhanh câu hỏi từ gợi ý.

### 3.6. Giao Diện Người Dùng (`AiAssistantControl.xaml`)
Nâng cấp khu vực Hỏi Đáp trong Sub-Tab 1:
1. **Thanh điều khiển phía trên:**
   - Badge hiển thị tên Sheet & vùng dữ liệu đang theo dõi (ví dụ: `📄 Data!A1:H80`).
   - Checkbox / Toggle: `[✓] Chỉ hỏi về vùng chọn`.
   - Nút `[🗑️ Xóa lịch sử / Chat mới]`.
2. **Dải nút Gợi ý câu hỏi nhanh (Prompt Pills):**
   - `📊 Tóm tắt bảng`
   - `🔍 Tìm ô bất thường / trống`
   - `🔝 Top 5 giá trị lớn nhất`
   - `💡 Gợi ý công thức`
3. **Khu vực hiển thị tin nhắn (Chat Message List - ScrollViewer):**
   - **Tin nhắn User:** Canh phải, nền xanh dương / Slate tối, bo góc kiểu hội thoại hiện đại.
   - **Tin nhắn AI Assistant:** Canh trái, nền card mềm mại (Trắng/Dark slate), viền phân cách, hỗ trợ:
     - Văn bản định dạng đẹp (bullet, icon).
     - Dải badge liên kết ô tính: `[🎯 A5]` `[🎯 C10:C25]` — click vào nhảy ô tức thì.
     - Khối đề xuất công thức kèm nút: `[🛠️ Chèn vào ô Excel]`.
     - Nút nhỏ sao chép câu trả lời.
4. **Hộp nhập liệu & Nút gửi phía dưới:**
   - Multi-line TextBox hỗ trợ gõ phím `Enter` (hoặc `Ctrl+Enter`) để gửi nhanh.
   - Nút gửi hình máy bay giấy `[➤ Gửi]`.
   - Hiệu ứng tải (Loading animation/spinner) khi AI đang phân tích dữ liệu.

---

## 4. Kế Hoạch Bản Quyền & Bản Địa Hóa (Localization)
Cập nhật đầy đủ chuỗi giao diện cho 3 ngôn ngữ: `vi.json`, `en.json`, `ja.json`:
* `Ai_ChatSheetTitle`: "Trò Chuyện Với Bảng Tính" / "Chat With Sheet" / "シートとチャット"
* `Ai_ChatSheetPlaceholder`: "Hỏi bất kỳ điều gì về dữ liệu bảng tính này..."
* `Ai_ChatScopeSelection`: "Chỉ vùng chọn"
* `Ai_ChatScopeAll`: "Toàn bộ Sheet"
* `Ai_PillSummarize`: "📊 Tóm tắt bảng dữ liệu"
* `Ai_PillAnomalies`: "🔍 Tìm giá trị bất thường"
* `Ai_PillTopRank`: "🔝 Top 5 giá trị cao nhất"
* `Ai_PillSuggestFormula`: "💡 Gợi ý công thức phân tích"
* `Ai_ReferencedCells`: "Vị trí dữ liệu liên quan:"
* `Ai_BtnInsertFormulaShort`: "Chèn vào ô"

---

## 5. Kế Hoạch Kiểm Thử & Xác Minh (Verification Plan)
1. **Kiểm tra trích xuất dữ liệu Sheet:**
   - Sheet rỗng hoặc ô đơn lẻ: Xác minh không gây lỗi COM exception, thông báo rõ ràng cho người dùng.
   - Sheet có bảng dữ liệu lớn (> 1000 dòng): Xác minh việc cắt giới hạn (maxRows = 100) an toàn, phản hồi dưới 3 giây.
   - Vùng chọn tùy ý (Non-contiguous hoặc Range đơn): Trích xuất chính xác cấu trúc bảng.
2. **Kiểm tra tính năng Nhảy ô (Cell Jump Link):**
   - Click vào badge `[A5]` $\rightarrow$ Excel nhảy đúng ô A5 và highlight.
   - Click vào badge dạng vùng `[B2:D10]` $\rightarrow$ Excel chọn toàn bộ vùng B2:D10.
   - Click vào badge có tên sheet `[Sheet2!C3]` $\rightarrow$ Excel chuyển sang Sheet2 và chọn C3.
3. **Kiểm tra Multi-turn Context:**
   - Hỏi câu 1: "Cột Doanh thu có tổng là bao nhiêu?"
   - Hỏi câu 2: "Ai là người có doanh thu cao nhất?"
   - Xác minh AI hiểu đại từ thay thế và ngữ cảnh của câu trước.
4. **Kiểm tra Giao diện Dark/Light Theme:**
   - Chuyển đổi theme qua nút mặt trăng/mặt trời, xác minh độ tương phản của bong bóng tin nhắn, màu chữ, nút bấm.
