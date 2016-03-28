using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public interface ISortersAndFilters
    {
        ITemplateDefinition TemplateDefinition { get; }
        List<IFilterDefinition> Filters { get; }
        List<ISorterDefinition> Sorters { get; }
        Type ResultType { get; }
        bool IsActive {get;}

        object Execute(IEnumerable<object> param);
        void Add(IFilterDefinition filterElement);
        void Remove(IFilterDefinition filterElement);
    }
}
