using System;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public interface ISorterDefinition
    {
        ITemplateDefinition TemplateDefinition { get; }
        IBindingDefinition BindingDefinition { get; }
        bool Descending  { get; }
        bool CaseSensitive { get; }
        object Sort(object source);
        Type ResultType  { get; }
    }
}
