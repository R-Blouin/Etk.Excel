namespace Etk.BindingTemplates.Context
{
    using System;
    using System.Collections.Generic;
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.BindingTemplates.Views;
    using Etk.SortAndFilter;

    public interface IBindingContext : IDisposable
    {
        /// <summary>Sorters and filters defined outside of the templates</summary>
        ISortersAndFilters ExternalSortsAndFilters { get; }

        ITemplateView Owner { get; }
        ITemplateDefinition TemplateDefinition { get; }
        IBindingContextElement Parent { get; }

        ///// <summary></summary>
        //IEnumerable<IBindingContextElement> Elements { get; }
        ///// <summary>Elements to render on the UI. Contains the element of the list 'Elements' filered and ordered by 'SortersAndFilters'</summary>
        //IEnumerable<IBindingContextElement> ElementsToRender { get; }

        /// <summary>Data source from which Elements and ElementsToRender are created</summary>
        object DataSource { get; }

        IBindingContextPart Header { get; }
        IBindingContextPart Body { get; }
        IBindingContextPart Footer { get; }

        List<IFilterDefinition> TemplatedFilters { get; }
    }
}
