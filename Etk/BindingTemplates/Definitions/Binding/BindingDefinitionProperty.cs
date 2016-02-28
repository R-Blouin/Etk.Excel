namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using Etk.BindingTemplates.Convertors;
    using Etk.Excel.UI.Log;

    class BindingDefinitionProperty : BindingDefinition
    {
        private ILogger log = Logger.Instance;

        private PropertyInfo PropertyInfo
        { get; set; }

        private MethodInfo GetMethod
        { get; set; } 

        private MethodInfo SetMethod
        { get; set; }

        #region .ctors and factories
        protected BindingDefinitionProperty(BindingDefinitionDescription bindingDefinitionDescription) : base(bindingDefinitionDescription)
        {}

        /// <summary>
        /// List<PropertyInfo> propertyInfos:  To detect the 'new' properties
        /// </summary>
        static public BindingDefinitionProperty CreateInstance(PropertyInfo propertyInfo, BindingDefinitionDescription definitionDescription)
        {
            if (propertyInfo == null)
                return null;
            
            if (definitionDescription == null)
                definitionDescription = new BindingDefinitionDescription() { Name = propertyInfo.Name };

            definitionDescription.IsReadOnly = definitionDescription.IsReadOnly || !propertyInfo.CanWrite || propertyInfo.GetSetMethod() == null;

            if (string.IsNullOrEmpty(definitionDescription.Description))
            {
                object[] descriptions = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), true);
                if(descriptions != null && descriptions.Any())
                    definitionDescription.Description = (descriptions[0] as DescriptionAttribute).Description;
            }

            BindingDefinitionProperty definition = new BindingDefinitionProperty(definitionDescription) {
                                                                                     BindingType = propertyInfo.PropertyType,
                                                                                     BindingTypeIsGeneric = propertyInfo.PropertyType.IsGenericType,
                                                                                     PropertyInfo = propertyInfo,
                                                                                     GetMethod = propertyInfo.GetGetMethod(),
                                                                                     SetMethod = propertyInfo.GetSetMethod()};
            if (definition.BindingTypeIsGeneric)
            {
                definition.BindingGenericType = definition.BindingType.GetGenericArguments()[0];
                definition.BindingGenericTypeDefinition = definition.BindingType.GetGenericTypeDefinition();
            }
            definition.CanNotify = !definition.IsACollection;

            definition.ManageCollectionStatus();
            definition.ManageEnumAndNullable();

            return definition;
        }
        #endregion

        #region override 'BindingDefinition' methods
        override public object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null && GetMethod != null)
                {
                    object ret = GetMethod.Invoke(GetMethod.IsStatic ? null : dataSource, null);
                    if (ret != null && ret is Enum)
                        return (ret as Enum).ToString();
                    else
                        return ret;
                }
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
        /// </summary>
        /// <param name="datasource">the Binding datasource to update.</param>
        /// <param name="data">the data to update the datasource.</param>
        /// <returns>
        /// If the BindingDefinition to update is readonly, then return the currently loaded value
        /// Else return the value passed as a parameter.
        /// </returns>
        override public object UpdateDataSource(object datasource, object data)
        {
            try
            {
                if (datasource == null)
                    return null;
 
                if (! IsReadOnly)
                {
                    Type type = BindingType;
                    if (data == null)
                        SetMethod.Invoke(SetMethod.IsStatic ? null : datasource, new object[] { (type.IsValueType ? Activator.CreateInstance(type) : null) });
                    else
                    {
                        data = SpecificConvertors.TryConvert(this, data);
                        SetMethod.Invoke(SetMethod.IsStatic ? null : datasource, new object[] { data });
                    }
                }
                return ResolveBinding(datasource);
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "'UpdateDataSource' failed for BindingExpression '{0}', value '{1}': {2}", BindingExpression, data == null ? string.Empty: data.ToString(), ex.Message);
                return ResolveBinding(datasource);
            }
        }

        override public IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            INotifyPropertyChanged notifyPropertyChanged = dataSource as INotifyPropertyChanged;
            return notifyPropertyChanged == null ? null : new INotifyPropertyChanged[] { notifyPropertyChanged };
        }
        
        override public bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return source == dataSource && BindingExpression.Equals(args.PropertyName);
        }
        #endregion
    }
}