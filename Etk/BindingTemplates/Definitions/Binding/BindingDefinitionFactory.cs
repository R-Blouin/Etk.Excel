using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.BindingTemplates.Definitions.Binding
{
    /// <summary> Binding definition factory</summary>
    public static class BindingDefinitionFactory
    {
        #region public static methods
        /// <summary> Create a binding definition for a given <see cref="BindingDefinitionDescription"/> owned by a given <see cref="FilterOwner"/> </summary>
        /// <param name="templateDefinition">The <see cref="FilterOwner"/> that owned the <see cref="BindingDefinitionDescription"/></param>
        /// <param name="bindingDefinitionDescription">the given <see cref="BindingDefinitionDescription"/></param>
        /// <returns>The newly created Binding definition or an exception is an error occurs</returns>
        public static IBindingDefinition CreateInstances(TemplateDefinition templateDefinition, BindingDefinitionDescription bindingDefinitionDescription)
        {
            IBindingDefinition ret = null;
            if (templateDefinition != null && bindingDefinitionDescription != null)
            {
                if(bindingDefinitionDescription.IsConst)
                    ret = BindingDefinitionConstante.CreateInstance(bindingDefinitionDescription);
                else
                {
                    Type type = templateDefinition.MainBindingDefinition?.BindingType;
                    ret = CreateInstance(type, bindingDefinitionDescription);
                }            
            }
            return ret;
        }

        /// <summary> Create a list of binding definition for a given list of <see cref="BindingDefinitionDescription"/> owned by a given <see cref="FilterOwner"/> </summary>
        /// <param name="templateDefinition">The <see cref="FilterOwner"/> that owned the <see cref="BindingDefinitionDescription"/></param>
        /// <param name="bindingDefinitionDescription">the given list of <see cref="BindingDefinitionDescription"/></param>
        /// <returns>The newly created Binding definition or an exception is an error occurs</returns>
        internal static List<IBindingDefinition> CreateInstances(Type type, List<BindingDefinitionDescription> definitionDescriptions)
        {
            List<IBindingDefinition> bindingDefinitions = new List<IBindingDefinition>();
            if (definitionDescriptions != null)
            {
                foreach (BindingDefinitionDescription definitionDescription in definitionDescriptions)
                {
                    IBindingDefinition bindingDefinition = CreateInstance(type, definitionDescription);
                    if (bindingDefinition != null)
                        bindingDefinitions.Add(bindingDefinition);
                }
            }
            return bindingDefinitions;
        }

        /// <summary> Create a binding definition for a given <see cref="BindingDefinitionDescription"/> of a given <see cref="Type"/> </summary>
        /// <param name="sourceType">The Type on which the '<see cref="BindingDefinitionDescription"/>' is based</param>
        /// <param name="definitionDescription">The given <see cref="Type"/></param>
        /// <returns>The newly created Binding definition or an exception is an error occurs</returns>
        internal static IBindingDefinition CreateInstance(Type sourceType, BindingDefinitionDescription definitionDescription)
        {
            if (string.IsNullOrEmpty(definitionDescription?.BindingExpression))
                return null;

            try
            {
                /// Composite
                /////////////
                if (definitionDescription.BindingExpression.StartsWith("{") && definitionDescription.BindingExpression.EndsWith("}"))
                    return BindingDefinitionComposite.CreateInstance(sourceType, definitionDescription);

                /// Hierarchical
                ////////////////
                if (definitionDescription.BindingExpression.Contains("."))
                    return BindingDefinitionHierarchical.CreateInstance(sourceType, definitionDescription);

                if (sourceType == null)
                    return BindingDefinitionOptional.CreateInstance(definitionDescription);

                //// keyword
                ////////////
                //if (definition == null && BindingDefinitionKeyWord.KeyWords.Contains(bindingName))
                //    definition = BindingDefinitionKeyWord.CreateInstances(bindingName);


                /// Properties
                //////////////
                List<PropertyInfo> propertyInfos = (from pi in sourceType.GetProperties()
                                                    where pi.Name.Equals(definitionDescription.BindingExpression) && pi.GetGetMethod() != null && pi.GetGetMethod().IsPublic
                                                    select pi).ToList();

                if (propertyInfos != null && propertyInfos.Count > 0)
                {
                    PropertyInfo propertyInfo;
                    if (propertyInfos.Count == 1)
                        propertyInfo = propertyInfos[0];
                    else // To take the keuword 'new' into account
                    {
                        propertyInfo = propertyInfos.FirstOrDefault(pi => { MethodInfo mi = pi.GetGetMethod();
                                                                            bool isNew = (mi.Attributes & MethodAttributes.HideBySig) == MethodAttributes.HideBySig && mi.DeclaringType.Equals(sourceType);
                                                                            return isNew; });
                    }
                    return BindingDefinitionProperty.CreateInstance(propertyInfo, definitionDescription);
                }

                /// Fields
                //////////
                FieldInfo fieldInfo = sourceType.GetFields().FirstOrDefault(fi => fi.Name.Equals(definitionDescription.BindingExpression) && fi.IsPublic);
                if (fieldInfo != null)
                    return BindingDefinitionField.CreateInstance(fieldInfo, definitionDescription);

                /// Methods
                ///////////
                MethodInfo methodInfo = sourceType.GetMethods().FirstOrDefault(mi => mi.Name.Equals(definitionDescription.BindingExpression) && mi.IsPublic);
                if (methodInfo != null)
                    return BindingDefinitionMethod.CreateInstance(methodInfo);
 
                return BindingDefinitionOptional.CreateInstance(definitionDescription);

            }
            catch (Exception ex)
            {
                throw new BindingTemplateException($"Cannot create 'BindingDefinition' between '{sourceType.Name}' and '{definitionDescription.BindingExpression}'", ex);
            }
        }

        /// <summary> Create a Binding definition for a given 'System.Reflection.PropertyInfo'</summary>
        /// <param name="propertyInfo">The given 'System.Reflection.PropertyInfo'</param>
        /// <returns>The newly created Bindiçng definition or an exception is an error occurs</returns>
        public static IBindingDefinition CreateInstance(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
                return null;

            return BindingDefinitionProperty.CreateInstance(propertyInfo, null);
        }
        #endregion
    }
}
