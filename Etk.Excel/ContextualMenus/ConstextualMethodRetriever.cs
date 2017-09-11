using System;
using System.Reflection;
using Etk.Tools.Reflection;

namespace Etk.Excel.ContextualMenus
{
    static class ConstextualMethodRetriever
    {
        public static MethodInfo RetrieveContextualMethodInfo(Type mainBindingDefinitionType, string methodName)
        {
            try
            {
                return TypeHelpers.GetMethod(mainBindingDefinitionType, methodName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Method '{methodName}' not resolved:{ex.Message}");
            }
        }

        public static MethodInfo RetrieveContextualMethodInfo(string methodName)
        {
            try
            {
                return TypeHelpers.GetMethod(null, methodName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Method '{methodName}' not resolved:{ex.Message}");
            }
        }
    }
}
