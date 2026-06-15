using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;

namespace PDFly.Converters;

/// <summary>true → Visible, false → Collapsed (and inverse if ConverterParameter is "Invert").</summary>
public sealed class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, string language)
    {
        bool b = value is bool x && x;
        if (parameter is string s && s.Equals("Invert", StringComparison.OrdinalIgnoreCase))
            b = !b;
        return b ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, string language)
        => throw new NotSupportedException();
}
