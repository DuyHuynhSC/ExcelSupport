using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelSupport.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelSupport.ViewModels
{
    public class AiAssistantViewModel : ViewModelBase
    {
        private int _selectedSubTab = 0; // 0: Dịch thuật, 1: Sinh công thức, 2: Gỡ lỗi & Hỏi đáp
        private bool _isBusy;
        private string _statusMessage = string.Empty;

        // --- Translation Properties ---
        private bool _writeToAdjacentColumn = false;
        private string _translationSummary = string.Empty;
        private int _translatedCellCount = 0;

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

        public bool WriteToAdjacentColumn
        {
            get => _writeToAdjacentColumn;
            set => SetProperty(ref _writeToAdjacentColumn, value);
        }

        public string TranslationSummary
        {
            get => _translationSummary;
            private set => SetProperty(ref _translationSummary, value);
        }

        public int TranslatedCellCount
        {
            get => _translatedCellCount;
            private set => SetProperty(ref _translatedCellCount, value);
        }

        public string FormulaPrompt
        {
            get => _formulaPrompt;
            set => SetProperty(ref _formulaPrompt, value);
        }

        public string FormulaResponse
        {
            get => _formulaResponse;
            private set => SetProperty(ref _formulaResponse, value);
        }

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
            private set => SetProperty(ref _cellInspectorSummary, value);
        }

        public string ChatPrompt
        {
            get => _chatPrompt;
            set => SetProperty(ref _chatPrompt, value);
        }

        public string ChatResponse
        {
            get => _chatResponse;
            private set => SetProperty(ref _chatResponse, value);
        }

        // --- Commands ---
        public ICommand TranslateJaToViCommand { get; }
        public ICommand TranslateViToJaCommand { get; }
        public ICommand GenerateFormulaCommand { get; }
        public ICommand InsertFormulaToExcelCommand { get; }
        public ICommand CopyFormulaCommand { get; }
        public ICommand ReadActiveCellCommand { get; }
        public ICommand DebugActiveCellCommand { get; }
        public ICommand SendChatCommand { get; }
        public ICommand ClearFormulaCommand { get; }
        public ICommand ClearChatCommand { get; }

        public AiAssistantViewModel()
        {
            TranslateJaToViCommand = new RelayCommand(async _ => await ExecuteTranslateSelectionAsync(isJaToVi: true), _ => !IsBusy);
            TranslateViToJaCommand = new RelayCommand(async _ => await ExecuteTranslateSelectionAsync(isJaToVi: false), _ => !IsBusy);
            GenerateFormulaCommand = new RelayCommand(async _ => await ExecuteGenerateFormulaAsync(), _ => !IsBusy && !string.IsNullOrWhiteSpace(FormulaPrompt));
            InsertFormulaToExcelCommand = new RelayCommand(_ => ExecuteInsertFormula());
            CopyFormulaCommand = new RelayCommand(_ => ExecuteCopyFormula());
            ReadActiveCellCommand = new RelayCommand(_ => ExecuteReadActiveCell());
            DebugActiveCellCommand = new RelayCommand(async _ => await ExecuteDebugActiveCellAsync(), _ => !IsBusy);
            SendChatCommand = new RelayCommand(async _ => await ExecuteSendChatAsync(), _ => !IsBusy && !string.IsNullOrWhiteSpace(ChatPrompt));
            ClearFormulaCommand = new RelayCommand(_ => ExecuteClearFormula());
            ClearChatCommand = new RelayCommand(_ => ExecuteClearChat());
        }

        #region Translation Logic (Japanese <-> Vietnamese)

        private async Task ExecuteTranslateSelectionAsync(bool isJaToVi)
        {
            var addIn = AddInEvents.Instance;
            if (addIn == null) return;

            var items = addIn.GetSelectedCellsText(maxCells: 300);
            if (items == null || items.Count == 0)
            {
                TranslationSummary = "⚠️ Vui lòng quét chọn các ô có chứa chữ trên Excel trước khi bấm dịch.";
                return;
            }

            IsBusy = true;
            string dirLabel = isJaToVi ? "Nhật ➔ Việt" : "Việt ➔ Nhật";
            StatusMessage = $"Đang dịch {items.Count} ô ({dirLabel})... ⏳";
            TranslationSummary = string.Empty;

            try
            {
                var config = AiConfigManager.Current;
                string srcLang = isJaToVi ? "tiếng Nhật" : "tiếng Việt";
                string tgtLang = isJaToVi ? "tiếng Việt" : "tiếng Nhật";

                // Build input JSON array
                var inputList = items.Select((item, idx) => new { id = idx, text = item.OriginalText }).ToList();
                string inputJson = JsonConvert.SerializeObject(inputList);

                string systemPrompt = $"Bạn là chuyên gia biên dịch ngôn ngữ công sở chuyên nghiệp giữa {srcLang} và {tgtLang} trong môi trường Excel/Doanh nghiệp. " +
                                      $"Hãy dịch từng mục trong mảng JSON được cung cấp sang {tgtLang}. " +
                                      $"Yêu cầu nghiêm ngặt: Trả về duy nhất một chuỗi JSON hợp lệ theo định dạng: " +
                                      $"[{{\"id\": 0, \"trans\": \"bản dịch\"}}, ...] không kèm bất kỳ lời giải thích hay markdown nào khác.";

                string userPrompt = $"Dịch mảng sau sang {tgtLang}:\n{inputJson}";

                string aiReply = await Task.Run(() => OpenAiClientService.SendChatAsync(config, userPrompt, systemPrompt));

                // Parse AI JSON response
                var transMap = ParseTranslationResponse(aiReply);

                for (int i = 0; i < items.Count; i++)
                {
                    if (transMap.TryGetValue(i, out string? translated) && !string.IsNullOrWhiteSpace(translated))
                    {
                        items[i].TranslatedText = translated;
                    }
                    else
                    {
                        items[i].TranslatedText = items[i].OriginalText;
                    }
                }

                // Write back to Excel
                bool ok = addIn.WriteTranslatedCells(items, WriteToAdjacentColumn);
                if (ok)
                {
                    string targetDesc = WriteToAdjacentColumn ? "ghi vào cột bên cạnh" : "ghi đè trực tiếp";
                    TranslationSummary = $"✅ Đã dịch thành công {items.Count} ô ({dirLabel}) và {targetDesc}!";
                    TranslatedCellCount = items.Count;
                }
                else
                {
                    TranslationSummary = "❌ Không thể ghi kết quả dịch vào Excel. Vui lòng kiểm tra quyền chỉnh sửa bảng tính.";
                }
            }
            catch (Exception ex)
            {
                TranslationSummary = $"❌ Lỗi dịch thuật: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                StatusMessage = string.Empty;
            }
        }

        private static Dictionary<int, string> ParseTranslationResponse(string response)
        {
            var result = new Dictionary<int, string>();
            if (string.IsNullOrWhiteSpace(response)) return result;

            try
            {
                string cleaned = response.Trim();
                int startIdx = cleaned.IndexOf('[');
                int endIdx = cleaned.LastIndexOf(']');
                if (startIdx >= 0 && endIdx > startIdx)
                {
                    cleaned = cleaned.Substring(startIdx, endIdx - startIdx + 1);
                }

                var array = JArray.Parse(cleaned);
                foreach (var token in array)
                {
                    if (token is JObject obj)
                    {
                        int id = obj["id"]?.Value<int>() ?? -1;
                        string trans = obj["trans"]?.ToString() ?? obj["translation"]?.ToString() ?? obj["text"]?.ToString() ?? "";
                        if (id >= 0 && !string.IsNullOrEmpty(trans))
                        {
                            result[id] = trans.Trim();
                        }
                    }
                }
            }
            catch
            {
                // Fallback nếu AI trả dạng từng dòng
                var lines = response.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    line = Regex.Replace(line, @"^\d+[\.\:\-\)]\s*", "");
                    if (!string.IsNullOrEmpty(line))
                    {
                        result[i] = line;
                    }
                }
            }
            return result;
        }

        #endregion

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
                string systemPrompt = "Bạn là chuyên gia Excel hàng đầu thế giới. Người dùng sẽ yêu cầu bạn tạo một công thức Excel từ mô tả tiếng Việt.\n" +
                                      "Yêu cầu:\n" +
                                      "1. Đưa ra công thức Excel CHUẨN XÁC NHẤT bắt đầu bằng dấu '=' và đặt bên trong khối code: ```excel\n=CÔNG_THỨC\n```\n" +
                                      "2. Ngay bên dưới khối code, hãy giải thích ngắn gọn, dễ hiểu bằng tiếng Việt về cách hoạt động của từng tham số trong công thức.\n" +
                                      "3. Ưu tiên các hàm hiện đại như XLOOKUP, SUMIFS, FILTER, UNIQUE, TEXTJOIN nếu phù hợp.";

                string response = await Task.Run(() => OpenAiClientService.SendChatAsync(config, FormulaPrompt, systemPrompt));

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

            // 2. Tìm dòng bất kỳ bắt đầu bằng '=' có chứa hàm Excel
            var lineMatch = Regex.Match(response, @"(?m)^\s*(=[A-Z_]+(?:\.[A-Z_]+)?\(.*?\))\s*$", RegexOptions.Multiline);
            if (lineMatch.Success)
            {
                return lineMatch.Groups[1].Value.Trim();
            }

            // 3. Tìm bất kỳ biểu thức =FUNCTION(...)
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

            try
            {
                var config = AiConfigManager.Current;
                string prompt = $"Phân tích và gỡ lỗi ô Excel sau:\n" +
                               $"- Sheet: {ActiveCell.SheetName}\n" +
                               $"- Tọa độ ô: {ActiveCell.CellAddress}\n" +
                               $"- Công thức: {ActiveCell.Formula}\n" +
                               $"- Giá trị hiển thị: {ActiveCell.Value}\n" +
                               $"- Lỗi: {(ActiveCell.HasError ? ActiveCell.ErrorText : "Không có")}\n\n" +
                               $"Hãy giải thích nguyên nhân gây ra lỗi (nếu có) và đưa ra công thức sửa lại hoàn chỉnh.";

                string systemPrompt = "Bạn là chuyên gia gỡ lỗi công thức Excel. Hãy đưa ra nguyên nhân lỗi rõ ràng và công thức khắc phục chuẩn xác bằng tiếng Việt.";

                string reply = await Task.Run(() => OpenAiClientService.SendChatAsync(config, prompt, systemPrompt));
                ChatResponse = reply;
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
            StatusMessage = "AI đang trả lời... ⏳";
            ChatResponse = string.Empty;

            try
            {
                var config = AiConfigManager.Current;
                string systemPrompt = "Bạn là trợ lý AI chuyên nghiệp về Excel, bảng tính và phân tích dữ liệu. Hãy trả lời câu hỏi của người dùng ngắn gọn, chính xác, dễ hiểu bằng tiếng Việt.";
                string reply = await Task.Run(() => OpenAiClientService.SendChatAsync(config, ChatPrompt, systemPrompt));
                ChatResponse = reply;
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

        private void ExecuteClearChat()
        {
            ChatPrompt = string.Empty;
            ChatResponse = string.Empty;
            StatusMessage = string.Empty;
        }

        #endregion
    }
}
