using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context
{
    public interface IBindingContextPart : IDisposable
    {
        IBindingContext ParentContext { get; }
        ISortersAndFilters ExternalSorterAndFilter  { get; }

        ITemplateDefinitionPart TemplateDefinitionPart { get; }

        IEnumerable<IBindingContextElement> Elements { get;}
        IEnumerable<IBindingContextElement> ElementsToRender{ get;}
    }
}
