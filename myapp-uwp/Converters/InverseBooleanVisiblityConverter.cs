using System;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Data;

namespace myapp_uwp.Converters
{
    public class InverseBooleanToVisibilityConverter : IValueConverter
    {
        public InverseBooleanToVisibilityConverter()
        {
        }

        public object Convert(object value, Type targetType, object parameter, string language)
        {
            if (value is bool && !(bool)value)
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            return (value is Visibility && (Visibility)value == Visibility.Visible);
        }
    }
}
