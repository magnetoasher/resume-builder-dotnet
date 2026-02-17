using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using ResumeBuilder.ViewModels;

namespace ResumeBuilder.Converters;

public class StepToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is AppStep step && parameter is string stepName && Enum.TryParse<AppStep>(stepName, out var target))
        {
            return step == target ? Visibility.Visible : Visibility.Collapsed;
        }

        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
