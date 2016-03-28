using System;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Convertors
{
    static class SpecificConvertors
    {
        public static object TryConvert(IBindingDefinition bindingDefinition, object data)
        {
            if (data == null)
                return null;

            if (bindingDefinition.BindingType.Equals(data.GetType()) || bindingDefinition.BindingType == typeof(object))
                return data;

            if (bindingDefinition.IsNullable && bindingDefinition.BindingGenericType.Equals(data.GetType()))
                return data;

            if (bindingDefinition.IsEnum)
            {
                string ret = data is string ? (string) data : data.ToString();
                Type type = bindingDefinition.IsNullable ? bindingDefinition.BindingGenericType : bindingDefinition.BindingType;
                return Enum.Parse(type, ret, true);
            }

            if (data.GetType() == typeof(string))
                return Convert.ChangeType(data, bindingDefinition.BindingType);

            if (bindingDefinition.BindingType.Equals(typeof(DateTime)) || (bindingDefinition.IsNullable && bindingDefinition.BindingGenericType.Equals(typeof(DateTime))))
            {
                if (data is double)
                    return DateTime.FromOADate((double) data);
                else if (data is string)
                {
                    //@@ Manage local
                    return DateTime.Parse(data as string);
                }
                else
                    return Convert.ChangeType(data, typeof(DateTime));
            }


            if (bindingDefinition.BindingType == typeof(bool) || (bindingDefinition.BindingTypeIsGeneric && bindingDefinition.BindingGenericType == typeof(bool)))
                return ToBoolean(bindingDefinition.BindingType, data);

            return Convert.ChangeType(data, bindingDefinition.BindingType);
        }

        private static object ToBoolean(Type type, object obj)
        {
            string objAsString;
            if (obj is string)
                objAsString = obj as string;
            else
                objAsString = obj.ToString();

            bool b;
            if (bool.TryParse(objAsString, out b))
                obj = b;
            else
            {
                objAsString = objAsString.Trim().ToUpper();
                if (objAsString.Equals("T") || objAsString.Equals("1"))
                    obj = true;
                else if (objAsString.Equals("F") || objAsString.Equals("0"))
                    obj = false;
            }
            return obj;
        }
    }
}
