using System;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.ModelManagement.DataAccessors;
using System.Reflection;

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
        public string Name => TemplateOption.Name;

        /// <summary> Implements <see cref="ITemplateDefinition.Description"/> </summary> 
        public string Description => !string.IsNullOrEmpty(TemplateOption.Description) ? TemplateOption.Description : Name;

        /// <summary> Implements <see cref="ITemplateDefinition.Orientation"/> </summary> 
        public Orientation Orientation => TemplateOption.Orientation;

        /// <summary> Implements <see cref="ITemplateDefinition.MainBindingDefinition"/> </summary> 
        public IBindingDefinition MainBindingDefinition
        { get; }

        /// <summary> Implements <see cref="ITemplateDefinition.DataAccessor"/> </summary> 
        public IDataAccessor DataAccessor => TemplateOption.DataAccessor;

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

        public bool CanSort => TemplateOption.CanSort;

        public bool AddBorder => TemplateOption.AddBorder;
        #endregion

        #region .ctors
        protected TemplateDefinition(TemplateOption templateOption)
        {
            if(templateOption == null)
                throw new ArgumentException("The template option cannot be null");

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

            BindingType = BindingType.CreateInstance(this);

            ((TemplateDefinitionPart) Header)?.Init();
            ((TemplateDefinitionPart) Body)?.Init();
            ((TemplateDefinitionPart) Footer)?.Init();
        }
        #endregion
    }
}