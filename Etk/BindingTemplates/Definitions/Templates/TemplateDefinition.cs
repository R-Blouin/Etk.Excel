using System;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.ModelManagement.DataAccessors;

namespace Etk.BindingTemplates.Definitions.Templates
{
    /// <summary>
    /// Template definition
    /// </summary>
    public class TemplateDefinition : ITemplateDefinition
    {
        #region attributes and properties
        /// <summary> Template options</summary>
        public TemplateOption TemplateOption
        { get; protected set; }

        /// <summary> Implements <see cref="ITemplateDefinition.Name"/> </summary> 
        public string Name
        { get { return TemplateOption.Name; } }

        /// <summary> Implements <see cref="ITemplateDefinition.Description"/> </summary> 
        public string Description
        { get { return !string.IsNullOrEmpty(TemplateOption.Description) ? TemplateOption.Description : Name; } }

        /// <summary> Implements <see cref="ITemplateDefinition.Orientation"/> </summary> 
        public Orientation Orientation
        { get { return TemplateOption.Orientation; } }

        /// <summary> Implements <see cref="ITemplateDefinition.MainBindingDefinition"/> </summary> 
        public IBindingDefinition MainBindingDefinition
        { get; private set; }

        /// <summary> Implements <see cref="ITemplateDefinition.DataAccessor"/> </summary> 
        public IDataAccessor DataAccessor
        { get { return TemplateOption.DataAccessor; } }

        /// <summary> Implements <see cref="ITemplateDefinition.Header"/> </summary> 
        public ITemplateDefinitionPart Header
        { get; protected set; }

        /// <summary> Implements <see cref="ITemplateDefinition.Body"/> </summary> 
        public ITemplateDefinitionPart Body
        { get; protected set; }

        /// <summary> Implements <see cref="ITemplateDefinition.Footer"/> </summary> 
        public ITemplateDefinitionPart Footer
        { get; protected set; }

        /// <summary> Implements <see cref="ITemplateDefinition.BindingType"/> </summary> 
        public BindingType BindingType
        { get; protected set; }

        public bool CanSort
        { get { return TemplateOption.CanSort;} }

        public bool AddBorder
        { get { return TemplateOption.AddBorder; } }
        #endregion

        #region .ctors
        protected TemplateDefinition(TemplateOption templateOption)
        {
            if(templateOption == null)
                throw new ArgumentException("The template option cannot be null");
            //@@if (string.IsNullOrEmpty(templateOption.Name))
            //    throw new ArgumentException("The template 'Name' cannot be null or empty");

            TemplateOption = templateOption;
            //&&HasLinkedTemplates = false;
            //BindingParts = new List<IDefinitionPart>();
            //LinkedTemplates = new List<ILinkedTemplateDefinition>();
            //BindingDefinitions = new List<IBindingDefinition>();
            MainBindingDefinition = templateOption.MainBindingDefinition;
        }

        protected TemplateDefinition(string name, TemplateDefinition parent)
        {
            if (parent == null)
                throw new ArgumentException("The parent template dataAccessor cannot be null");

            TemplateOption = parent.TemplateOption;
            //&&HasLinkedTemplates = false;
            //BindingParts = new List<IDefinitionPart>();
            //LinkedTemplates = new List<ILinkedTemplateDefinition>();
            //BindingDefinitions = new List<IBindingDefinition>();

            MainBindingDefinition = parent.MainBindingDefinition;
        }
        #endregion

        #region protected metrhods
        protected virtual void Init(ITemplateDefinitionPart header, ITemplateDefinitionPart body, ITemplateDefinitionPart footer)
        {
            Header = header;
            Body = body;
            Footer = footer;

            BindingType = Etk.BindingTemplates.Definitions.Binding.BindingType.CreateInstance(this);

            if (Header != null)
                ((TemplateDefinitionPart) Header).Init();
            if (Body != null)
                ((TemplateDefinitionPart) Body).Init();
            if (Footer != null)
                ((TemplateDefinitionPart) Footer).Init();
        }
        #endregion
    }
}