using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using ExcelSupport.Models;

namespace ExcelSupport.Helpers
{
    public static class RichTextHelper
    {
        public static readonly DependencyProperty FormattedItemProperty =
            DependencyProperty.RegisterAttached(
                "FormattedItem",
                typeof(CellTextItem),
                typeof(RichTextHelper),
                new PropertyMetadata(null, OnFormattedItemChanged));

        public static CellTextItem? GetFormattedItem(DependencyObject obj)
        {
            return (CellTextItem?)obj.GetValue(FormattedItemProperty);
        }

        public static void SetFormattedItem(DependencyObject obj, CellTextItem? value)
        {
            obj.SetValue(FormattedItemProperty, value);
        }

        private static void OnFormattedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock textBlock) return;

            textBlock.Inlines.Clear();

            if (e.NewValue is not CellTextItem item) return;

            if (item.FormattedRuns != null && item.FormattedRuns.Count > 0)
            {
                foreach (var runModel in item.FormattedRuns)
                {
                    if (string.IsNullOrEmpty(runModel.Text)) continue;

                    var run = new Run(runModel.Text);
                    if (runModel.IsStrikethrough)
                    {
                        run.TextDecorations = TextDecorations.Strikethrough;
                    }
                    if (runModel.IsBold)
                    {
                        run.FontWeight = FontWeights.Bold;
                    }
                    if (runModel.IsItalic)
                    {
                        run.FontStyle = FontStyles.Italic;
                    }
                    if (!string.IsNullOrEmpty(runModel.ColorHex))
                    {
                        try
                        {
                            var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(runModel.ColorHex);
                            run.Foreground = new SolidColorBrush(color);
                        }
                        catch { }
                    }

                    textBlock.Inlines.Add(run);
                }
            }
            else
            {
                textBlock.Text = item.OriginalText ?? string.Empty;
            }
        }

        public static readonly DependencyProperty TranslatedItemProperty =
            DependencyProperty.RegisterAttached(
                "TranslatedItem",
                typeof(CellTextItem),
                typeof(RichTextHelper),
                new PropertyMetadata(null, OnTranslatedItemChanged));

        public static CellTextItem? GetTranslatedItem(DependencyObject obj)
        {
            return (CellTextItem?)obj.GetValue(TranslatedItemProperty);
        }

        public static void SetTranslatedItem(DependencyObject obj, CellTextItem? value)
        {
            obj.SetValue(TranslatedItemProperty, value);
        }

        private static void OnTranslatedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock textBlock) return;

            textBlock.Inlines.Clear();

            if (e.NewValue is not CellTextItem item) return;

            if (item.TranslatedRuns != null && item.TranslatedRuns.Count > 0)
            {
                foreach (var runModel in item.TranslatedRuns)
                {
                    if (string.IsNullOrEmpty(runModel.Text)) continue;

                    var run = new Run(runModel.Text);
                    if (runModel.IsStrikethrough)
                    {
                        run.TextDecorations = TextDecorations.Strikethrough;
                    }
                    if (runModel.IsBold)
                    {
                        run.FontWeight = FontWeights.Bold;
                    }
                    if (runModel.IsItalic)
                    {
                        run.FontStyle = FontStyles.Italic;
                    }
                    if (!string.IsNullOrEmpty(runModel.ColorHex))
                    {
                        try
                        {
                            var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(runModel.ColorHex);
                            run.Foreground = new SolidColorBrush(color);
                        }
                        catch { }
                    }

                    textBlock.Inlines.Add(run);
                }
            }
            else
            {
                textBlock.Text = item.TranslatedText ?? string.Empty;
            }
        }
    }
}
