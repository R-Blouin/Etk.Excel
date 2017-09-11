using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionMethod : BindingDefinition
    {
        private readonly ILogger log = Logger.Instance;

        public MethodInfo MethodInfo
        { get; protected set; }

        #region override 'BindingDefinition' methods
        public override object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null && MethodInfo != null)
                    return MethodInfo.Invoke(MethodInfo.IsStatic ? null : dataSource, null);
                return null;
            }
            catch (Exception ex)
            {
                throw new BindingTemplateException($"Cannot Resolve the 'Binding' for the BindingExpression '{BindingExpression}'. {ex.Message}");
            }
        }

        /// <summary>
        /// Update the datasource.
        /// If the BindingDefinition to update is readonly, then return the currently laoded value
        /// Else return the value passed as a parameter. 
        /// </summary>
        public override object UpdateDataSource(object dataSource, object data)
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

        public static BindingDefinitionMethod CreateInstance(MethodInfo methodInfo)
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