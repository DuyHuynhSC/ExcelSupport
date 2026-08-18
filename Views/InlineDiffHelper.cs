using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using ExcelSupport.Models;
using MediaColor = System.Windows.Media.Color;

namespace ExcelSupport.Views
{
    public static class InlineDiffHelper
    {
        public static readonly DependencyProperty DiffSegmentsProperty =
            DependencyProperty.RegisterAttached(
                "DiffSegments",
                typeof(IEnumerable<TextDiffSegment>),
                typeof(InlineDiffHelper),
                new PropertyMetadata(null, OnDiffSegmentsChanged));

        public static IEnumerable<TextDiffSegment> GetDiffSegments(DependencyObject obj) =>
            (IEnumerable<TextDiffSegment>)obj.GetValue(DiffSegmentsProperty);

        public static void SetDiffSegments(DependencyObject obj, IEnumerable<TextDiffSegment> value) =>
            obj.SetValue(DiffSegmentsProperty, value);

        private static void OnDiffSegmentsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock tb)
            {
                tb.Inlines.Clear();
                if (e.NewValue is IEnumerable<TextDiffSegment> segments)
                {
                    foreach (var seg in segments)
                    {
                        var run = new Run(seg.Text);
                        if (seg.IsDiff)
                        {
                            run.Foreground = new SolidColorBrush(MediaColor.FromRgb(239, 68, 68)); // Đỏ nổi bật #EF4444
                            run.FontWeight = FontWeights.Bold;
                            run.Background = new SolidColorBrush(MediaColor.FromArgb(45, 239, 68, 68)); // Nền đỏ mờ
                        }
                        tb.Inlines.Add(run);
                    }
                }
            }
        }
    }
}
