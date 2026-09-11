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

        private static readonly DependencyProperty FormattedHandlerProperty =
            DependencyProperty.RegisterAttached(
                "FormattedHandler",
                typeof(System.ComponentModel.PropertyChangedEventHandler),
                typeof(RichTextHelper),
                new PropertyMetadata(null));

        private static void OnFormattedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock textBlock) return;

            if (e.OldValue is CellTextItem oldItem)
            {
                if (textBlock.GetValue(FormattedHandlerProperty) is System.ComponentModel.PropertyChangedEventHandler oldHandler)
                {
                    oldItem.PropertyChanged -= oldHandler;
                    textBlock.ClearValue(FormattedHandlerProperty);
                }
            }

            if (e.NewValue is CellTextItem newItem)
            {
                System.ComponentModel.PropertyChangedEventHandler handler = (s, ev) =>
                {
                    if (ev.PropertyName == nameof(CellTextItem.FormattedRuns) ||
                        ev.PropertyName == nameof(CellTextItem.OriginalText) ||
                        string.IsNullOrEmpty(ev.PropertyName))
                    {
                        RenderFormattedInlines(textBlock, newItem);
                    }
                };
                textBlock.SetValue(FormattedHandlerProperty, handler);
                newItem.PropertyChanged += handler;

                RenderFormattedInlines(textBlock, newItem);
            }
            else
            {
                textBlock.Inlines.Clear();
            }
        }

        private static void RenderFormattedInlines(TextBlock textBlock, CellTextItem item)
        {
            textBlock.Inlines.Clear();

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

        private static readonly DependencyProperty TranslatedHandlerProperty =
            DependencyProperty.RegisterAttached(
                "TranslatedHandler",
                typeof(System.ComponentModel.PropertyChangedEventHandler),
                typeof(RichTextHelper),
                new PropertyMetadata(null));

        private static void OnTranslatedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock textBlock) return;

            if (e.OldValue is CellTextItem oldItem)
            {
                if (textBlock.GetValue(TranslatedHandlerProperty) is System.ComponentModel.PropertyChangedEventHandler oldHandler)
                {
                    oldItem.PropertyChanged -= oldHandler;
                    textBlock.ClearValue(TranslatedHandlerProperty);
                }
            }

            if (e.NewValue is CellTextItem newItem)
            {
                System.ComponentModel.PropertyChangedEventHandler handler = (s, ev) =>
                {
                    if (ev.PropertyName == nameof(CellTextItem.TranslatedRuns) ||
                        ev.PropertyName == nameof(CellTextItem.TranslatedText) ||
                        string.IsNullOrEmpty(ev.PropertyName))
                    {
                        RenderTranslatedInlines(textBlock, newItem);
                    }
                };
                textBlock.SetValue(TranslatedHandlerProperty, handler);
                newItem.PropertyChanged += handler;

                RenderTranslatedInlines(textBlock, newItem);
            }
            else
            {
                textBlock.Inlines.Clear();
            }
        }

        private static void RenderTranslatedInlines(TextBlock textBlock, CellTextItem item)
        {
            textBlock.Inlines.Clear();

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
