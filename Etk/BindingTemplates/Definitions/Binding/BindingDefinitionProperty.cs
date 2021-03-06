﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Convertors;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionProperty : BindingDefinition
    {
        private readonly ILogger log = Logger.Instance;

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
        public static BindingDefinitionProperty CreateInstance(PropertyInfo propertyInfo, BindingDefinitionDescription definitionDescription)
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
        public override object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null && GetMethod != null)
                {
                    object ret = GetMethod.Invoke(GetMethod.IsStatic ? null : dataSource, null);
                    //if (ret != null && ret is Enum)
                    //    return (ret as Enum).ToString();
                    return ret;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new BindingTemplateException($"Cannot Resolve the 'Binding' for the BindingExpression '{BindingExpression}'. {ex.Message}");
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
        public override object UpdateDataSource(object datasource, object data)
        {
            try
            {
                if (datasource == null)
                    return null;
 
                if (! IsReadOnly)
                {
                    Type type = BindingType;
                    if (data == null)
                        SetMethod.Invoke(SetMethod.IsStatic ? null : datasource, new object[] { type.IsValueType ? Activator.CreateInstance(type) : null });
                    else
                    {
                        data = SpecificConvertors.TryConvert(this, data);
                        SetMethod.Invoke(SetMethod.IsStatic ? null : datasource, new object[] { data });
                    }
                }
                object value = ResolveBinding(datasource);
                return value is Enum ? ((Enum) value).ToString() : value;
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "'UpdateDataSource' failed for BindingExpression '{0}', value '{1}': {2}", BindingExpression, data == null ? string.Empty: data.ToString(), ex.Message);
                return ResolveBinding(datasource);
            }
        }

        public override IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            INotifyPropertyChanged notifyPropertyChanged = dataSource as INotifyPropertyChanged;
            return notifyPropertyChanged == null ? null : new [] { notifyPropertyChanged };
        }
        
        public override bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return source == dataSource && BindingExpression.Equals(args.PropertyName);
        }
        #endregion
    }
}