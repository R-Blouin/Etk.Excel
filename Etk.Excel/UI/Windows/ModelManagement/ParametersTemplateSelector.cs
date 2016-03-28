using System.Windows.Controls;

namespace Etk.Excel.UI.Windows.ModelManagement
{
    public class ParametersTemplateSelector : DataTemplateSelector
    {
        //public DataTemplate DefaultTemplate
        //{ get; set; }

        //public DataTemplate EnumTemplate
        //{ get; set; }

        //public override DataTemplate SelectTemplate(object obj, DependencyObject container)
        //{
        //    if (obj == null)
        //        return null;

        //    AccessorParameter accessorParameter = obj as AccessorParameter;
        //    if (accessorParameter == null || accessorParameter.ParameterInfos ==null)
        //        return null;

        //    FrameworkElement element = container as FrameworkElement;

        //    Type type = accessorParameter.ParameterInfos.ParameterType;
        //    if (type.Name.Equals("DateTime") || type.IsGenericType && type.GetGenericArguments()[0].Name.Equals("DateTime"))
        //        return element.FindResource("DateTimeTemplate") as DataTemplate;

        //    return element.FindResource("DefaultTemplate") as DataTemplate;
        //}
    }
}
