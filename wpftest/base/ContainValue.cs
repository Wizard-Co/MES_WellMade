using System;
using System.Globalization;
using System.Windows.Data;

namespace WizMes_WellMade
{
    public class ContainValue : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value.ToString().Contains(flag);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public string flag { get; set; }
    }
}
