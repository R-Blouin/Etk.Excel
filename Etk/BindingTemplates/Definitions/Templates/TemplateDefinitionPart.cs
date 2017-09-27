using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public class TemplateDefinitionPart : ITemplateDefinitionPart 
    {
        public ITemplateDefinition Parent
        {get; protected set;}

        public List<IBindingDefinition> BindingDefinitions
        { get; protected set; }

        public List<ILinkedTemplateDefinition> LinkedTemplates
        { get; protected set; }

        public List<IDefinitionPart> BindingParts
        { get; }

        public List<BindingFilterDefinition> FilterDefinitions
        { get; }
        
        public bool HasLinkedTemplates
        { get; protected set; }

        public TemplateDefinitionPartType PartType
        { get; }

        #region .ctors
        public TemplateDefinitionPart(TemplateDefinitionPartType partType)
        {
            HasLinkedTemplates = false;
            BindingParts = new List<IDefinitionPart>();
            LinkedTemplates = new List<ILinkedTemplateDefinition>();
            BindingDefinitions = new List<IBindingDefinition>();
            FilterDefinitions = new List<BindingFilterDefinition>();

            PartType = partType;
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

        public virtual void Init()
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
