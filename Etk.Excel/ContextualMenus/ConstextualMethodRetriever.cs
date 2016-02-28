namespace Etk.Excel.ContextualMenus
{
    using System;
    using System.Reflection;
    using Etk.Excel.UI.Reflection;

    static class ConstextualMethodRetriever
    {
        static public MethodInfo RetrieveContextualMethodInfo(Type mainBindingDefinitionType, string methodName)
        {
            try
            {
                return TypeHelpers.GetMethod(mainBindingDefinitionType, methodName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(string.Format("Method '{0}' not resolved:{0}", methodName, ex.Message), ex);
            }
        }

        static public MethodInfo RetrieveContextualMethodInfo(string methodName)
        {
            try
            {
                return TypeHelpers.GetMethod(null, methodName);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(string.Format("Method '{0}' not resolved:{0}", methodName, ex.Message), ex);
            }
        }
    }
}
