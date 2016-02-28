namespace Etk.Excel.UI.Reflection
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    /// <summary> Internal use</summary>
    public static class TypeConvertor
    {
        public static object ConvertObject(Type type, object value)
        {
            try
            {
                if (value == null)
                    return null;

                if(type.Equals(value.GetType()) || type == typeof(object))
                    return value;
                if (type.IsGenericType && type.GetGenericArguments()[0].Equals(value.GetType()))
                    return value;
                if (value.GetType() == typeof(string))
                    return Convert.ChangeType(value, type);

                if (type.Equals(typeof(DateTime)))
                {
                    if (value is double)
                        return DateTime.FromOADate((double)value);
                    else if (value is string)
                    {
                        //@@ Gérer local
                        return DateTime.Parse(value as string);
                    }
                }

                bool isValueACollection = ! value.GetType().Equals(typeof(string))
                                          && value.GetType().GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));
                if (isValueACollection)
                {
                    Type genericType = type.GetGenericArguments()[0];
                    MethodInfo convertCollection = typeof(TypeConvertor).GetMethod("ConvertCollection").MakeGenericMethod(genericType);
                    return convertCollection.Invoke(null, new object[] { genericType, value });
                }
                else
                {
                    bool isTypeACollection = type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)                                                                                               || i == typeof(System.Collections.IEnumerable));
                    if (isTypeACollection)
                    {
                        Type genericType = type.GetGenericArguments()[0];
                        MethodInfo convertToCollection = typeof(TypeConvertor).GetMethod("ConvertToCollection").MakeGenericMethod(genericType);
                        return convertToCollection.Invoke(null, new object[] { genericType, value });
                    }
                    else
                        return Convert.ChangeType(value, type);
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("'ConvertObject' failed. Can't convert '{0}' to UnderlyingType '{1}'. {2}", value == null ? string.Empty : value.ToString(), type.ToString(), ex.Message));
            }
        }

        public static IEnumerable<T> ConvertCollection<T>(Type type, IEnumerable values)
        {
            List<T> list = new List<T>();
            foreach (object o in values)
            {
                T ret = (T) Convert.ChangeType(o, type);
                list.Add(ret);
            }
            return list;
        }

        public static IEnumerable<T> ConvertToCollection<T>(Type type, object value)
        {
            List<T> list = new List<T>();

            T ret = (T) Convert.ChangeType(value, type);
            list.Add(ret);
            return list;
        }
    }
}
