namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using Etk.Excel.UI.Log;

    class BindingDefinitionMethod : BindingDefinition
    {
        private ILogger log = Logger.Instance;

        public MethodInfo MethodInfo
        { get; protected set; }

        #region override 'BindingDefinition' methods
        override public object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null && MethodInfo != null)
                    return MethodInfo.Invoke(MethodInfo.IsStatic ? null : dataSource, null);
                return null;
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot Resolve the 'Binding' for the BindingExpression '{0}'. {1}", BindingExpression, ex.Message);
                throw new BindingTemplateException(message, ex);
            }
        }

        /// <summary>
        /// Update the datasource.
        /// If the BindingDefinition to update is readonly, then return the currently laoded value
        /// Else return the value passed as a parameter. 
        /// </summary>
        override public object UpdateDataSource(object dataSource, object data)
        {
           try
            {
                return ResolveBinding(dataSource);
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "Cannot Resolve 'UpdateDataSource' for the BindingExpression '{0}'. {1}", BindingExpression, ex.Message);
                return ResolveBinding(dataSource);
            }
        }
        #endregion
 
        #region .ctors and factories
        private BindingDefinitionMethod(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        { }

        static public BindingDefinitionMethod CreateInstance(MethodInfo methodInfo)
        {
            if (methodInfo == null)
                return null;

            BindingDefinitionDescription definitionDescription = new BindingDefinitionDescription() { BindingExpression = methodInfo.Name, IsReadOnly = true };
            if (string.IsNullOrEmpty(definitionDescription.Description))
            {
                object[] descriptions = methodInfo.GetCustomAttributes(typeof(DescriptionAttribute), true);
                if (descriptions != null && descriptions.Any())
                    definitionDescription.Description = (descriptions[0] as DescriptionAttribute).Description;
            }

            BindingDefinitionMethod definition = new BindingDefinitionMethod(definitionDescription) { BindingType = methodInfo.ReturnType, 
                                                                                                      MethodInfo = methodInfo};
            definition.ManageCollectionStatus();
            definition.ManageEnumAndNullable();

            return definition;
        }
        #endregion
    }
}