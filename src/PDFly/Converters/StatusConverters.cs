using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using PDFly.Models;
using Windows.UI;

namespace PDFly.Converters;

/// <summary>Maps <see cref="ConversionStatus"/> to a Fluent / MDL2 glyph code point.</summary>
public sealed class StatusToGlyphConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
        => value is ConversionStatus s
            ? s switch
            {
                ConversionStatus.Pending    => "", // Document
                ConversionStatus.Converting => "", // Sync
                ConversionStatus.Done       => "", // CheckMark
                ConversionStatus.Skipped    => "", // Warning
                ConversionStatus.Failed     => "", // ErrorBadge
                _ => "",
            }
            : "";

    public object ConvertBack(object value, Type targetType, object parameter, string language)
        => throw new NotSupportedException();
}

/// <summary>Maps <see cref="ConversionStatus"/> to a colour brush for the row glyph.</summary>
public sealed class StatusToBrushConverter : IValueConverter
{
    private static readonly SolidColorBrush Pending = Make(Color.FromArgb(0xFF, 0x8A, 0x8A, 0x8A));
    private static readonly SolidColorBrush Done    = Make(Color.FromArgb(0xFF, 0x2B, 0xA8, 0x4A));
    private static readonly SolidColorBrush Skipped = Make(Color.FromArgb(0xFF, 0xC9, 0x88, 0x1B));
    private static readonly SolidColorBrush Failed  = Make(Color.FromArgb(0xFF, 0xE0, 0x47, 0x3C));

    public object Convert(object value, Type targetType, object parameter, string language)
    {
        if (value is not ConversionStatus s) return Pending;
        if (s == ConversionStatus.Converting)
            return (Application.Current?.Resources?["AccentFillColorDefaultBrush"] as Brush) ?? Pending;
        return s switch
        {
            ConversionStatus.Done    => Done,
            ConversionStatus.Skipped => Skipped,
            ConversionStatus.Failed  => Failed,
            _ => Pending,
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
        => throw new NotSupportedException();

    private static SolidColorBrush Make(Color c) => new(c);
}
