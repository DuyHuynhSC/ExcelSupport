# Chat With Sheet (Task Pane) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Triển khai tính năng Chat With Sheet trong Task Pane của ExcelSupport, cho phép người dùng hỏi đáp tương tác trực tiếp với dữ liệu bảng tính hiện tại bằng tiếng Việt tự nhiên kèm link nhảy nhanh tới ô tính (`[A1]`, `[B2:D10]`) và nút 1-Click chèn công thức vào Excel.

**Architecture:** Bổ sung `ChatSheetMessageItem` model; mở rộng `AddInEvents.AiHelpers.cs` với chức năng bóc tách ngữ cảnh Sheet thông minh (`GetSheetChatContext`) và hàm nhảy ô Excel (`NavigateToCell`); nâng cấp `OpenAiClientService.cs` hỗ trợ hội thoại nhiều lượt (`SendChatMessagesAsync`); cập nhật `AiAssistantViewModel.cs` quản lý luồng tin nhắn và tương tác cell link; thiết kế lại giao diện hỏi đáp trong `AiAssistantControl.xaml` thành giao diện chat hiện đại hỗ trợ dark/light theme và phím tắt gửi tin nhắn.

**Tech Stack:** C# .NET Framework 4.8 / WPF, Excel-DNA 1.8.0, Microsoft.Office.Interop.Excel, Newtonsoft.Json.

**Spec:** [docs/superpowers/specs/2026-09-11-chat-with-sheet-design.md](../../specs/2026-09-11-chat-with-sheet-design.md)

## Global Constraints
- Target Framework: .NET Framework 4.8
- UI Framework: WPF (Windows Presentation Foundation) với MVVM pattern
- An toàn luồng COM Excel: Luôn sử dụng `ExcelAsyncUtil.QueueAsMacro` khi gọi thao tác COM tương tác với Excel từ luồng background hoặc UI event.
- Giải phóng COM Objects: Sử dụng `Marshal.ReleaseComObject` hoặc `ComScope` để tránh rò rỉ bộ nhớ COM.
- Bản địa hóa: Đầy đủ 3 ngôn ngữ (`vi.json`, `en.json`, `ja.json`).

---

### Task 1: Model Dữ Liệu `ChatSheetMessageItem`
**Files:**
- Create: `Models/ChatSheetMessageItem.cs`

**Interfaces:**
- Produces: `ExcelSupport.Models.ChatSheetMessageItem` kế thừa `ViewModelBase`.

- [ ] **Step 1: Tạo file `Models/ChatSheetMessageItem.cs`**
```csharp
using System;
using System.Collections.ObjectModel;
using ExcelSupport.ViewModels;

namespace ExcelSupport.Models
{
    public class ChatSheetMessageItem : ViewModelBase
    {
        private string _content = string.Empty;
        private string? _suggestedFormula;
        private bool _isUser;

        public string Id { get; set; } = Guid.NewGuid().ToString();

        public bool IsUser
        {
            get => _isUser;
            set => SetProperty(ref _isUser, value);
        }

        public string Content
        {
            get => _content;
            set => SetProperty(ref _content, value);
        }

        public DateTime Timestamp { get; set; } = DateTime.Now;

        public ObservableCollection<string> ReferencedCells { get; } = new ObservableCollection<string>();

        public bool HasReferencedCells => ReferencedCells.Count > 0;

        public string? SuggestedFormula
        {
            get => _suggestedFormula;
            set
            {
                if (SetProperty(ref _suggestedFormula, value))
                {
                    OnPropertyChanged(nameof(HasSuggestedFormula));
                }
            }
        }

        public bool HasSuggestedFormula => !string.IsNullOrWhiteSpace(SuggestedFormula);

        public void NotifyCellsChanged()
        {
            OnPropertyChanged(nameof(HasReferencedCells));
        }
    }
}
```

- [ ] **Step 2: Kiểm tra biên dịch**
Run: `dotnet build ExcelSupport.csproj`
Expected: Build succeeded.

- [ ] **Step 3: Commit**
```bash
git add Models/ChatSheetMessageItem.cs
git commit -m "feat(ai): add ChatSheetMessageItem model for chat with sheet"
```

---

### Task 2: Dịch Vụ AI Client Hỗ Trợ Hội Thoại Nhiều Lượt
**Files:**
- Modify: `Services/OpenAiClientService.cs`

**Interfaces:**
- Produces: `OpenAiClientService.SendChatMessagesAsync(AiConfig config, IEnumerable<(string role, string content)> messages, string? systemPrompt = null)`

- [ ] **Step 1: Thêm phương thức `SendChatMessagesAsync` vào `Services/OpenAiClientService.cs`**
Cung cấp khả năng gửi mảng hội thoại (`role` & `content`), tự động áp dụng `max_completion_tokens` hoặc `max_tokens` và cơ chế fallback.

```csharp
public static async Task<string> SendChatMessagesAsync(AiConfig config, IEnumerable<(string role, string content)> messages, string? systemPrompt = null)
{
    string baseUrl = NormalizeBaseUrl(config.BaseUrl);
    string endpoint = $"{baseUrl}/chat/completions";
    string model = string.IsNullOrWhiteSpace(config.ModelName) ? "qwen-3.6" : config.ModelName.Trim();

    var messagesArray = new JArray();
    if (!string.IsNullOrWhiteSpace(systemPrompt))
    {
        messagesArray.Add(new JObject { ["role"] = "system", ["content"] = systemPrompt });
    }

    foreach (var (role, content) in messages)
    {
        if (!string.IsNullOrWhiteSpace(content))
        {
            messagesArray.Add(new JObject { ["role"] = role, ["content"] = content });
        }
    }

    var payload = new JObject
    {
        ["model"] = model,
        ["messages"] = messagesArray
    };

    if (IsReasoningModel(model))
    {
        payload["max_completion_tokens"] = config.MaxTokens;
    }
    else
    {
        payload["max_tokens"] = config.MaxTokens;
        payload["temperature"] = config.Temperature;
    }

    using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(10, config.TimeoutSeconds))))
    {
        using (var response = await SendWithFallbackAsync(endpoint, config.ApiKey, payload, cts.Token))
        {
            string responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                string errDetail = ParseErrorMessage(responseBody) ?? response.ReasonPhrase ?? "Lỗi API";
                throw new InvalidOperationException($"Lỗi từ máy chủ AI (HTTP {(int)response.StatusCode}): {errDetail}");
            }

            return ExtractAssistantReply(responseBody) ?? string.Empty;
        }
    }
}
```

- [ ] **Step 2: Kiểm tra biên dịch**
Run: `dotnet build ExcelSupport.csproj`
Expected: Build succeeded.

- [ ] **Step 3: Commit**
```bash
git add Services/OpenAiClientService.cs
git commit -m "feat(ai): add SendChatMessagesAsync multi-turn support to OpenAiClientService"
```

---

### Task 3: Excel COM Helpers Cho Ngữ Cảnh Bảng Tính & Nhảy Ô
**Files:**
- Modify: `Host/AddInEvents.AiHelpers.cs`

**Interfaces:**
- Produces:
  - `AddInEvents.Instance.GetSheetChatContext(bool selectionOnly, int maxRows = 100, int maxCols = 25)`
  - `AddInEvents.Instance.NavigateToCell(string cellAddressOrRange)`

- [ ] **Step 1: Triển khai `GetSheetChatContext` và `NavigateToCell` trong `Host/AddInEvents.AiHelpers.cs`**
Hỗ trợ:
- Đọc vùng chọn (Selection) mảng 2 chiều COM siêu tốc, format Markdown bảng.
- Nếu không chọn vùng hoặc `selectionOnly == false`: đọc `UsedRange` (giới hạn 100 dòng, 25 cột), trích xuất headers + thông tin ActiveCell.
- Hàm `NavigateToCell`: Tách sheet name (nếu có dạng `'Sheet Name'!A1` hoặc `Sheet1!A1`), kích hoạt sheet, cuộn và chọn vùng bằng `app.Goto(range, true)`.

- [ ] **Step 2: Kiểm tra biên dịch**
Run: `dotnet build ExcelSupport.csproj`
Expected: Build succeeded.

- [ ] **Step 3: Commit**
```bash
git add Host/AddInEvents.AiHelpers.cs
git commit -m "feat(ai): add GetSheetChatContext and NavigateToCell helpers in AddInEvents"
```

---

### Task 4: Mở Rộng ViewModel `AiAssistantViewModel`
**Files:**
- Modify: `ViewModels/AiAssistantViewModel.cs`

**Interfaces:**
- Consumes: `ChatSheetMessageItem`, `OpenAiClientService.SendChatMessagesAsync`, `AddInEvents.GetSheetChatContext`, `AddInEvents.NavigateToCell`.
- Produces:
  - `ObservableCollection<ChatSheetMessageItem> SheetChatMessages`
  - `string SheetChatInput`
  - `bool IsSelectionOnly`
  - `ICommand SendSheetChatCommand`
  - `ICommand ClearSheetChatCommand`
  - `ICommand NavigateToCellCommand`
  - `ICommand ApplySuggestedFormulaCommand`
  - `ICommand QuickPromptCommand`

- [ ] **Step 1: Cập nhật `AiAssistantViewModel.cs`**
- Khởi tạo danh sách tin nhắn mẫu chào mừng đầu tiên từ AI.
- Hàm regex trích xuất các cell link từ câu trả lời: ví dụ `\[([A-Za-z0-9_]+![A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?|[A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?)\]` hoặc các mẫu ô `ô ([A-Z]+[0-9]+)`.
- Hàm `ExecuteSendSheetChatAsync`:
  1. Thêm user message vào `SheetChatMessages`.
  2. Thu thập ngữ cảnh từ `GetSheetChatContext(IsSelectionOnly)`.
  3. Lập danh sách messages (ngữ cảnh hệ thống + tối đa 6 lượt tin nhắn gần nhất).
  4. Gọi `SendChatMessagesAsync`.
  5. Phân tích câu trả lời, trích xuất cell references và suggested formula.
  6. Thêm assistant message vào `SheetChatMessages`.
- Lệnh `NavigateToCellCommand`: Chuyển tiếp tới `AddInEvents.Instance.NavigateToCell(address)`.
- Lệnh `ApplySuggestedFormulaCommand`: Chèn công thức vào ô active.
- Lệnh `QuickPromptCommand`: Gán prompt mẫu và kích hoạt gửi ngay.

- [ ] **Step 2: Kiểm tra biên dịch**
Run: `dotnet build ExcelSupport.csproj`
Expected: Build succeeded.

- [ ] **Step 3: Commit**
```bash
git add ViewModels/AiAssistantViewModel.cs
git commit -m "feat(ai): integrate Chat With Sheet multi-turn logic in AiAssistantViewModel"
```

---

### Task 5: Cập Nhật Giao Diện WPF & Bản Địa Hóa
**Files:**
- Modify: `Views/AiAssistantControl.xaml`
- Modify: `Views/AiAssistantControl.xaml.cs`
- Modify: `Languages/vi.json`
- Modify: `Languages/en.json`
- Modify: `Languages/ja.json`

- [ ] **Step 1: Cập nhật các file ngôn ngữ (`vi.json`, `en.json`, `ja.json`)**
Thêm các chuỗi giao diện cho Chat With Sheet.

- [ ] **Step 2: Cập nhật giao diện `AiAssistantControl.xaml`**
- Nâng cấp phần Hỏi Đáp trong Sub-Tab 1 thành giao diện Chat With Sheet hoàn chỉnh.
- Hỗ trợ cuộn tự động, bong bóng tin nhắn, badge link ô tính, dark/light theme triggers.

- [ ] **Step 3: Cập nhật code-behind `AiAssistantControl.xaml.cs`**
- Hỗ trợ phím tắt `Enter` để gửi tin nhắn (hoặc `Shift+Enter` để xuống dòng) trong ô chat.
- Tự động cuộn xuống đáy khi có tin nhắn mới trong `SheetChatScrollViewer`.

- [ ] **Step 4: Kiểm tra biên dịch toàn diện & Đóng gói Add-In**
Run: `dotnet build ExcelSupport.csproj`
Expected: 0 Warnings, 0 Errors, file XLL được build và đóng gói thành công.

- [ ] **Step 5: Commit**
```bash
git add Views/AiAssistantControl.xaml Views/AiAssistantControl.xaml.cs Languages/
git commit -m "feat(ui): complete Chat With Sheet UI in Task Pane with cell jump badges and localization"
```
