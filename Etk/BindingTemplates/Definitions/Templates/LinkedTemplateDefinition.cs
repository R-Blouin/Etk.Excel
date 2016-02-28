namespace Etk.BindingTemplates.Definitions.Templates
{
    using System;
    using Etk.BindingTemplates.Definitions.Binding;

    public class LinkedTemplateDefinition : ILinkedTemplateDefinition 
    {
        public BindingPartType PartType
        { get { return BindingPartType.LinkedTemplateDefinition; } }

        public ITemplateDefinition Parent
        { get; private set; }

        public ITemplateDefinition TemplateDefinition
        {get; private set;}

        public IBindingDefinition BindingDefinition
        { get; private set; }

        public LinkedTemplatePositioning Positioning
        { get; private set; }

        #region .ctors
        public LinkedTemplateDefinition(ITemplateDefinition parent, ITemplateDefinition templateDefinition, TemplateLink linkDefinition)
        {
            Parent = parent;
            TemplateDefinition = templateDefinition;
            Positioning = linkDefinition.Positioning;
            if (! string.IsNullOrEmpty(linkDefinition.With))
            {
                Type type = parent.MainBindingDefinition != null ? parent.MainBindingDefinition.BindingType : null;
                BindingDefinitionDescription definitionDescription = new BindingDefinitionDescription()
                {
                    BindingExpression = linkDefinition.With,
                    Description = linkDefinition.Description,
                    Name = linkDefinition.Name,
                    IsReadOnly = true
                };
                BindingDefinition = BindingDefinitionFactory.CreateInstance(type, definitionDescription);
            }
        }
        #endregion

        #region public methods
        public object ResolveBinding(object dataSource)
        {
            if (dataSource == null)
                return null;

            if (BindingDefinition == null)
                return dataSource;

            if (BindingDefinition.IsOptional)
                BindingDefinition = (BindingDefinition as BindingDefinitionOptional).CreateRealBindingDefinition(dataSource.GetType());
       
            return BindingDefinition.ResolveBinding(dataSource);
        }
        #endregion
    }
}
