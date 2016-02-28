namespace Etk.BindingTemplates.Definitions.Templates
{
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.Templates.Xml;
    using Etk.Excel.UI.Reflection;
    using ModelManagement.DataAccessors;
    using System;
    using System.Linq;
    using System.Reflection;

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

        public ExpanderMode ExpanderMode
        { get; set; }

        public bool HeaderAsExpander
        { get; set; }

        public IBindingDefinition ExpanderBindingDefinition
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
                Description = xmlTemplateOption.Description;
                Orientation = xmlTemplateOption.Orientation;
                ExpanderMode = xmlTemplateOption.ExpanderMode;
                HeaderAsExpander = xmlTemplateOption.HeaderAsExpander;
                AddBorder = xmlTemplateOption.AddBorder;
                CanSort = xmlTemplateOption.CanSort;

                SelectionChanged = xmlTemplateOption.SelectionChanged;
                ContextualMenu = xmlTemplateOption.ContextMenu;

                if (!string.IsNullOrEmpty(xmlTemplateOption.Decorator))
                    DecoratorIdent = xmlTemplateOption.Decorator;

                //if (string.IsNullOrEmpty(xmlTemplateOption.BindingType) && string.IsNullOrEmpty(xmlTemplateOption.BindingMethod))
                //    throw new ArgumentException("One of the two elements 'BindingType' and 'BindingMethod' must be defined");

                //if (!string.IsNullOrEmpty(xmlTemplateOption.BindingType) && !string.IsNullOrEmpty(xmlTemplateOption.BindingMethod))
                //    throw new ArgumentException("The 'BindingMethod' and 'BindingType' elements cannot be both defined");

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
                //        throw new EtkException(string.Format("Cannot resolve the DataAccessor '{0}':{1}", bindingMethod, ex.Message), ex);
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
                        throw new EtkException(string.Format("Cannot resolve 'BindingWith' '{0}':{1}", bindingWith, ex.Message), ex);
                    }
                }

                // Retrieve the Expander binding definition
                ///////////////////////////////////////////
                if (! string.IsNullOrEmpty(xmlTemplateOption.Expander))
                {
                    try
                    {
                        string property = xmlTemplateOption.Expander.Trim();
                        PropertyInfo propertyInfo = MainBindingDefinition.BindingType.GetProperties().FirstOrDefault(pi => pi.Name.Equals(property));
                        ExpanderBindingDefinition = BindingDefinitionFactory.CreateInstance(propertyInfo);
                    }
                    catch (Exception ex)
                    {
                        throw new EtkException(string.Format("Cannot resolve the expander binding definition from '{0}'. {1}", xmlTemplateOption.Expander, ex.Message), ex);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Resolve the template oprion failed:{0}", ex.Message), ex);
            }
        }
        #endregion

        #region pubolic methods
        #endregion
    }
}
