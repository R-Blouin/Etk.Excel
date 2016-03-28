using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public interface IFilterDefinition
    {
        ITemplateDefinition TemplateDefinition { get; }
        IBindingDefinition DefinitionToFilter { get; }
        string FilterExpression { get; }
    }
}
