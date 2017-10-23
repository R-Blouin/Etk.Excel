using System;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates.Xml;
using Etk.ModelManagement.DataAccessors;
using Etk.Tools.Reflection;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public class TemplateOption
    {
        #region attributes and properties
        public string Name
        { get; set; }

        public string Description
        { get; set; }

        public Orientation Orientation
        { get; set; }

        public IDataAccessor DataAccessor
        { get; set; }

        public IBindingDefinition MainBindingDefinition
        { get; set; }

        public ExpanderType ExpanderType
        { get; set; }

        public HeaderAsExpander HeaderAsExpander
        { get; set; }

        public string SelectionChanged
        { get; set; }

        public string ContextualMenu
        { get; set; }

        
        public bool CanSort
        { get; set; }

        public bool AddBorder
        { get; set; }

        public string DecoratorIdent
        { get; set; }
        #endregion

        #region .ctors and factories
        public TemplateOption()
        {}

        public TemplateOption(XmlTemplateOption xmlTemplateOption)
        {
            try
            {
                Name = xmlTemplateOption.Name;
                if (string.IsNullOrEmpty(xmlTemplateOption.Name))
                    throw new ArgumentException("The template 'Name' cannot be null or empty");

                Description = xmlTemplateOption.Description;
                Orientation = xmlTemplateOption.Orientation;
                ExpanderType = xmlTemplateOption.ExpanderType;
                HeaderAsExpander = xmlTemplateOption.HeaderAsExpander;
                AddBorder = xmlTemplateOption.AddBorder;

                SelectionChanged = xmlTemplateOption.SelectionChanged;
                ContextualMenu = xmlTemplateOption.ContextMenu;

                if (!string.IsNullOrEmpty(xmlTemplateOption.Decorator))
                    DecoratorIdent = xmlTemplateOption.Decorator;

                // Retrieve the 'MainBindingDefinition'
                ///////////////////////////////////////
                //if (!string.IsNullOrEmpty(xmlTemplateOption.BindingMethod))
                //{
                //    // Get the Data Source accessor method
                //    //////////////////////////////////////
                //    string bindingMethod = xmlTemplateOption.BindingMethod.Trim();
                //    try
                //    {
                //        DataAccessorInstanceType dataAccessorInstanceType = ModelManagement.DataAccessors.DataAccessor.AccessorInstanceTypeFrom(xmlTemplateOption.InstanceType);
                //        DataAccessor = ModelManagement.DataAccessors.DataAccessor.CreateInstance(bindingMethod, dataAccessorInstanceType, xmlTemplateOption.InstanceName);
                //        MainBindingDefinition = BindingDefinitionRoot.CreateInstance(DataAccessor.ReturnType);
                //    }
                //    catch (Exception ex)
                //    {
                //        throw new EtkException(string.Format("Cannot resolve the DataAccessor '{0}':{1}", bindingMethod, ex.Message));
                //    }
                //}
                //else
                if (!string.IsNullOrEmpty(xmlTemplateOption.BindingWith))
                {
                    // Get the main binding type
                    ////////////////////////////
                    string bindingWith = xmlTemplateOption.BindingWith.Trim();
                    try
                    {
                        Type type = TypeHelpers.GetType(bindingWith);
                        if (type == null)
                            throw new EtkException("Type not found");
                        MainBindingDefinition = BindingDefinitionRoot.CreateInstance(type);
                    }
                    catch (Exception ex)
                    {
                        throw new EtkException($"Cannot resolve 'BindingWith' '{bindingWith}':{ex.Message}");
                    }
                }

                if (xmlTemplateOption.CanSort.HasValue)
                {
                    if (xmlTemplateOption.CanSort.Value && MainBindingDefinition == null)
                        throw new EtkException("'CanSort' parameter can be set to 'true' only if the 'BindingWith' parameter is set.");

                    CanSort = xmlTemplateOption.CanSort.Value;
                }
                else
                    CanSort = MainBindingDefinition != null;
            }
            catch (Exception ex)
            {
                throw new EtkException($"Resolve the template option failed:{ex.Message}");
            }
        }
        #endregion
    }
}
