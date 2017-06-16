using System;
using System.ComponentModel;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Convertors
{
    static class SpecificConvertors
    {
        public static object TryConvert(IBindingDefinition bindingDefinition, object data)
        {
            if (data == null)
                return null;

            if (bindingDefinition.BindingType == data.GetType() || bindingDefinition.BindingType == typeof(object))
                return data;

            Type toConvertType = bindingDefinition.IsNullable ? bindingDefinition.BindingGenericType : bindingDefinition.BindingType;

            if (toConvertType == data.GetType())
                return data;

            if (toConvertType == typeof(DateTime))
                return ToDateTime(data);

            if (toConvertType == typeof(bool))
                return ToBoolean(data);

            if (toConvertType.IsEnum)
                return ToEnum(toConvertType, data);

            return Convert.ChangeType(data, toConvertType);
        }

        private static object ToBoolean(object obj)
        {
            string objAsString = obj as string ?? obj.ToString();

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

        private static object ToDateTime(object data)
        {
            if (data is double)
                return DateTime.FromOADate((double)data);
            if (data is string) //@@ Manage local
                return DateTime.Parse(data as string);
            return Convert.ChangeType(data, typeof(DateTime));           
        }

        private static object ToEnum(Type type, object data)
        {
            string ret = data is string ? (string) data : data.ToString();
            return Enum.Parse(type, ret, true);
        }
    }
}
