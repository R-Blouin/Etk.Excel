namespace Etk.SortAndFilter
{
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.Templates;
    
    public interface IFilterDefinition
    {
        ITemplateDefinition TemplateDefinition { get; }
        IBindingDefinition DefinitionToFilter { get; }
        string FilterExpression { get; }
    }
}
