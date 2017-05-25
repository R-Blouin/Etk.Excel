using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace Etk.Excel.UI.Windows.SortAndFilter.Converters
{
    [ValueConversion(typeof(object), typeof(int))]  
    class SelectionConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isSelected = (bool) value;
            return isSelected ? 2 : 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
