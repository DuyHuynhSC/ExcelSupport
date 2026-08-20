using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using ExcelSupport.Host;
using WpfMessageBox = System.Windows.MessageBox;

namespace ExcelSupport.Views
{
    public class ManualChapterItem
    {
        public string Title { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public int Index { get; set; }
    }

    public partial class UserManualDialog : Window
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(nameof(IsDarkTheme), typeof(bool), typeof(UserManualDialog),
                new PropertyMetadata(false, OnThemeChangedStatic));

        private static void OnThemeChangedStatic(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UserManualDialog dlg && dlg.ListChapters?.SelectedItem is ManualChapterItem item)
            {
                dlg.RenderChapter(item);
            }
        }

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private static UserManualDialog? _currentInstance;

        internal static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.IsDarkTheme = isDarkTheme;
                    _currentInstance.Activate();
                    return;
                }

                var addIn = AddInEvents.Instance;
                var app = addIn?.ExcelAppInstance;

                _currentInstance = new UserManualDialog
                {
                    IsDarkTheme = isDarkTheme
                };

                try
                {
                    if (app != null)
                    {
                        new System.Windows.Interop.WindowInteropHelper(_currentInstance).Owner = (IntPtr)app.Hwnd;
                    }
                }
                catch { }

                _currentInstance.Show();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi mở sách hướng dẫn sử dụng:\n{ex.Message}",
                                   "ExcelSupport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<ManualChapterItem> _allChapters = new List<ManualChapterItem>();
        private string _fullManualMarkdown = string.Empty;

        public UserManualDialog()
        {
            InitializeComponent();

            try
            {
                IsDarkTheme = AddInEvents.MainViewModel?.IsDarkTheme ?? false;
            }
            catch { }

            Loaded += OnDialogLoaded;
            Closed += (s, e) => _currentInstance = null;
        }

        private void OnDialogLoaded(object sender, RoutedEventArgs e)
        {
            LoadManualContent();
        }

        private void LoadManualContent()
        {
            try
            {
                // 1. Thử đọc từ file cục bộ trước nếu có
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string localFile = Path.Combine(baseDir, "USER_MANUAL.md");
                string devFile = @"D:\SourceCode\ExcelSupport\USER_MANUAL.md";

                if (File.Exists(localFile))
                {
                    _fullManualMarkdown = File.ReadAllText(localFile, System.Text.Encoding.UTF8);
                }
                else if (File.Exists(devFile))
                {
                    _fullManualMarkdown = File.ReadAllText(devFile, System.Text.Encoding.UTF8);
                }
                else
                {
                    // 2. Đọc từ EmbeddedResource
                    var asm = Assembly.GetExecutingAssembly();
                    string resourceName = asm.GetManifestResourceNames()
                        .FirstOrDefault(r => r.EndsWith("USER_MANUAL.md", StringComparison.OrdinalIgnoreCase)) ?? "";

                    if (!string.IsNullOrEmpty(resourceName))
                    {
                        using var stream = asm.GetManifestResourceStream(resourceName);
                        if (stream != null)
                        {
                            using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                            _fullManualMarkdown = reader.ReadToEnd();
                        }
                    }
                }

                if (string.IsNullOrEmpty(_fullManualMarkdown))
                {
                    _fullManualMarkdown = "# Hướng Dẫn Sử Dụng\nKhông tìm thấy file tài liệu `USER_MANUAL.md`.";
                }

                ParseChapters();
            }
            catch (Exception ex)
            {
                _fullManualMarkdown = $"# Lỗi Nạp Tài Liệu\n{ex.Message}";
                ParseChapters();
            }
        }

        private void ParseChapters()
        {
            _allChapters.Clear();

            var lines = _fullManualMarkdown.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var currentTitle = "0. Tổng Quan";
            var currentContent = new System.Text.StringBuilder();
            int idx = 0;

            foreach (var line in lines)
            {
                if (line.StartsWith("## ") && !line.StartsWith("### "))
                {
                    if (currentContent.Length > 0)
                    {
                        _allChapters.Add(new ManualChapterItem
                        {
                            Index = idx++,
                            Title = currentTitle,
                            Content = currentContent.ToString()
                        });
                        currentContent.Clear();
                    }
                    currentTitle = line.Substring(3).Trim();
                }
                currentContent.AppendLine(line);
            }

            if (currentContent.Length > 0)
            {
                _allChapters.Add(new ManualChapterItem
                {
                    Index = idx++,
                    Title = currentTitle,
                    Content = currentContent.ToString()
                });
            }

            ListChapters.ItemsSource = null;
            ListChapters.ItemsSource = _allChapters;
            ListChapters.DisplayMemberPath = nameof(ManualChapterItem.Title);

            if (_allChapters.Count > 0)
            {
                ListChapters.SelectedIndex = 0;
            }
        }

        private void OnSearchChapterChanged(object sender, TextChangedEventArgs e)
        {
            string q = TxtSearchChapter.Text.Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(q))
            {
                ListChapters.ItemsSource = _allChapters;
            }
            else
            {
                ListChapters.ItemsSource = _allChapters
                    .Where(c => c.Title.ToLowerInvariant().Contains(q) || c.Content.ToLowerInvariant().Contains(q))
                    .ToList();
            }
            ListChapters.DisplayMemberPath = nameof(ManualChapterItem.Title);
        }

        private void OnChapterSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListChapters.SelectedItem is ManualChapterItem item)
            {
                RenderChapter(item);
            }
        }

        private void RenderChapter(ManualChapterItem item)
        {
            FlowDoc.Blocks.Clear();

            var textColor = IsDarkTheme 
                ? (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#F8FAFC")! 
                : (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#0F172A")!;
            var subTextColor = IsDarkTheme 
                ? (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#94A3B8")! 
                : (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#475569")!;
            var headerColor = IsDarkTheme 
                ? (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#38BDF8")! 
                : (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#0369A1")!;
            var accentColor = (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#107C41")!;

            var lines = item.Content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                string tr = line.Trim();
                if (string.IsNullOrEmpty(tr)) continue;

                if (tr.StartsWith("# "))
                {
                    var p = new Paragraph(new Run(tr.Substring(2).Trim()))
                    {
                        FontSize = 18,
                        FontWeight = FontWeights.Bold,
                        Foreground = accentColor,
                        Margin = new Thickness(0, 8, 0, 8)
                    };
                    FlowDoc.Blocks.Add(p);
                }
                else if (tr.StartsWith("## "))
                {
                    var p = new Paragraph(new Run(tr.Substring(3).Trim()))
                    {
                        FontSize = 16,
                        FontWeight = FontWeights.Bold,
                        Foreground = headerColor,
                        Margin = new Thickness(0, 10, 0, 6)
                    };
                    FlowDoc.Blocks.Add(p);
                }
                else if (tr.StartsWith("### "))
                {
                    var p = new Paragraph(new Run(tr.Substring(4).Trim()))
                    {
                        FontSize = 14,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = headerColor,
                        Margin = new Thickness(0, 8, 0, 4)
                    };
                    FlowDoc.Blocks.Add(p);
                }
                else if (tr.StartsWith("---"))
                {
                    // Separator
                    var p = new Paragraph
                    {
                        BorderThickness = new Thickness(0, 0, 0, 1),
                        BorderBrush = IsDarkTheme 
                            ? (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#334155")! 
                            : (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#E2E8F0")!,
                        Margin = new Thickness(0, 6, 0, 6)
                    };
                    FlowDoc.Blocks.Add(p);
                }
                else
                {
                    var p = new Paragraph
                    {
                        FontSize = 12.5,
                        Foreground = textColor,
                        Margin = new Thickness(0, 2, 0, 3)
                    };

                    string cleanText = tr;
                    if (cleanText.StartsWith("* ") || cleanText.StartsWith("- "))
                    {
                        p.Margin = new Thickness(14, 2, 0, 3);
                        cleanText = "• " + cleanText.Substring(2);
                    }

                    // Format bold tags **text**
                    var parts = Regex.Split(cleanText, @"(\*\*.*?\*\*)");
                    foreach (var part in parts)
                    {
                        if (part.StartsWith("**") && part.EndsWith("**") && part.Length > 4)
                        {
                            p.Inlines.Add(new Bold(new Run(part.Substring(2, part.Length - 4))));
                        }
                        else
                        {
                            p.Inlines.Add(new Run(part));
                        }
                    }

                    FlowDoc.Blocks.Add(p);
                }
            }
        }

        private void OnOpenRawFileClick(object sender, RoutedEventArgs e)
        {
            try
            {
                string devFile = @"D:\SourceCode\ExcelSupport\USER_MANUAL.md";
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string localFile = Path.Combine(baseDir, "USER_MANUAL.md");

                string target = File.Exists(devFile) ? devFile : localFile;

                if (!File.Exists(target))
                {
                    File.WriteAllText(localFile, _fullManualMarkdown, System.Text.Encoding.UTF8);
                    target = localFile;
                }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = target,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Không thể mở file:\n{ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
