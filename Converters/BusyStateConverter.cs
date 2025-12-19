using System;
using System.Globalization;
using System.Windows.Data;

namespace OSEMAddIn.Converters
{
    internal sealed class BusyStateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var isBusy = value is bool flag && flag;
            return isBusy ? Properties.Resources.Extracting : Properties.Resources.Ready;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
