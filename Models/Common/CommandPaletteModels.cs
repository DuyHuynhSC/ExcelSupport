using System;
using System.Collections.Generic;
using ExcelSupport.Services;

namespace ExcelSupport.Models
{
    /// <summary>
    /// Đại diện cho một lệnh/tính năng trong Quick Command Palette
    /// </summary>
    public class PaletteCommandItem
    {
        public string Id { get; set; } = string.Empty;
        public string TitleKey { get; set; } = string.Empty;
        public string DescriptionKey { get; set; } = string.Empty;
        public string CategoryKey { get; set; } = string.Empty;
        public string IconEmoji { get; set; } = "⚡";
        public string ShortcutText { get; set; } = string.Empty;
        public List<string> Keywords { get; set; } = new List<string>();
        public Action? Action { get; set; }

        // Trạng thái hiển thị động
        public bool IsRecent { get; set; }
        public int MatchScore { get; set; }

        public string Title => !string.IsNullOrEmpty(TitleKey) 
            ? LocalizationService.Get(TitleKey) 
            : Id;

        public string Description => !string.IsNullOrEmpty(DescriptionKey) 
            ? LocalizationService.Get(DescriptionKey) 
            : string.Empty;

        public string Category => !string.IsNullOrEmpty(CategoryKey) 
            ? LocalizationService.Get(CategoryKey) 
            : LocalizationService.Get("grpQuickTools");
    }

    /// <summary>
    /// Thông tin lưu trữ lệnh được sử dụng gần đây
    /// </summary>
    public class RecentCommandEntry
    {
        public string CommandId { get; set; } = string.Empty;
        public DateTime LastUsedUtc { get; set; } = DateTime.UtcNow;
    }
}
