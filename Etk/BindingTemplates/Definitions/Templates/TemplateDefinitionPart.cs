namespace Etk.BindingTemplates.Definitions.Templates
{
    using System.Collections.Generic;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.SortSearchAndFilter;

    public class TemplateDefinitionPart : ITemplateDefinitionPart 
    {
        public ITemplateDefinition Parent
        {get; protected set;}

        public List<IBindingDefinition> BindingDefinitions
        { get; protected set; }

        public List<ILinkedTemplateDefinition> LinkedTemplates
        { get; protected set; }

        public List<IDefinitionPart> BindingParts
        { get; private set; }

        public List<BindingFilterDefinition> FilterDefinitions
        { get; private set; }
        
        public bool HasLinkedTemplates
        { get; protected set; }

        #region .ctors
        public TemplateDefinitionPart()
        {
            HasLinkedTemplates = false;
            BindingParts = new List<IDefinitionPart>();
            LinkedTemplates = new List<ILinkedTemplateDefinition>();
            BindingDefinitions = new List<IBindingDefinition>();
            FilterDefinitions = new List<BindingFilterDefinition>();
        }
        #endregion

        #region public methods
        public void AddLinkedTemplate(LinkedTemplateDefinition linkedTemplateDefinition)
        {
            BindingParts.Add(linkedTemplateDefinition);
            LinkedTemplates.Add(linkedTemplateDefinition);
            HasLinkedTemplates = true;
        }

        public void AddBindingDefinition(IBindingDefinition definition)
        {
            BindingParts.Add(definition);
            BindingDefinitions.Add(definition);
        }

        public void AddFilterDefinition(BindingFilterDefinition definition)
        {
            BindingParts.Add(definition);
            FilterDefinitions.Add(definition);
        }

        public void AddSearchDefinition(BindingSearchDefinition definition)
        {
            BindingParts.Add(definition);
        }

        virtual public void Init()
        {
            if (FilterDefinitions != null && FilterDefinitions.Count > 0)
            {
                foreach (BindingFilterDefinition bindingFilterDefinition in FilterDefinitions)
                    bindingFilterDefinition.Init();
            }
        }
        #endregion
    }
}
