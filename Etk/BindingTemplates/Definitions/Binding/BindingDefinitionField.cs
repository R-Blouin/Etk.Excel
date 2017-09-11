using System;
using System.Reflection;
using Etk.BindingTemplates.Convertors;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionField : BindingDefinition
    {
        private readonly ILogger log = Logger.Instance;

        private FieldInfo FieldInfo
        { get; set; }

        #region override 'BindingDefinition' methods
        public override object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null)
                {
                    object ret = FieldInfo.GetValue(FieldInfo.IsStatic ? null : dataSource);
                    //if (ret != null && ret is Enum)
                    //    return (ret as Enum).ToString();
                    return ret;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new BindingTemplateException(
                    $"Can't Resolve the 'Binding' for the BindingExpression '{BindingExpression}'. {ex.Message}");
            }
        }

        
        /// <summary>
        /// Update the datasource with the binding value.
        /// If the BindingDefinition to update is readonly, then return the currently loaded value
        /// Else return the value passed as a parameter. 
        /// </summary>
        public override object UpdateDataSource(object dataSource, object data)
        {
            try
            {
                if (dataSource == null)
                    return null;

                if (!IsReadOnly)
                {
                    Type type = BindingType;
                    if (data == null)
                        FieldInfo.SetValue(FieldInfo.IsStatic ? null : dataSource, type.IsValueType ? Activator.CreateInstance(type) : null);
                    else
                    {
                        data = SpecificConvertors.TryConvert(this, data);
                        FieldInfo.SetValue(FieldInfo.IsStatic ? null : dataSource, data);
                    }
                }
                return ResolveBinding(dataSource);
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "'UpdateDataSource' failed for BindingExpression '{0}', value '{1}': {2}", BindingExpression, data?.ToString() ?? string.Empty, ex.Message);
                return ResolveBinding(dataSource);
            }
        }
        #endregion

        #region .ctors
        private BindingDefinitionField(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        { }
        #endregion

        #region static public methods
        public static BindingDefinitionField CreateInstance(FieldInfo fieldInfo, BindingDefinitionDescription definitionDescription)
        {
            BindingDefinitionField definition = new BindingDefinitionField(definitionDescription) {
                                                                           BindingType = fieldInfo.FieldType,
                                                                           BindingTypeIsGeneric = fieldInfo.FieldType.IsGenericType,
                                                                           FieldInfo = fieldInfo};
            if (definition.BindingTypeIsGeneric)
            {
                definition.BindingGenericType = definition.BindingType.GetGenericArguments()[0];
                definition.BindingGenericTypeDefinition = definition.BindingType.GetGenericTypeDefinition();
            }
            definition.ManageCollectionStatus();
            definition.ManageEnumAndNullable();

            return definition;
        }
        #endregion
    }
}