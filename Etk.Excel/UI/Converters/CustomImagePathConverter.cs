namespace Etk.Excel.UI.Converters
{
    using System;
    using System.Globalization;
    using System.Windows.Data;

    public class CustomImagePathConverter : IValueConverter
    {
        #region IValueConverter Members
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return "../Images/" + GetImageName(value.ToString());
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return "";
        }
        #endregion

        private string GetImageName(string text)
        {
            string name = "";
            name = text.ToLower() + ".png";
            return name;
        }
    } 
}
