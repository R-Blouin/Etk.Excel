using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Views
{
    public interface ITemplateView : IDisposable
    {
        Guid Ident { get; }

        object GetDataSource();
        void SetDataSource(object dataSource);
        //void ApplyFilter();

        IBindingContext BindingContext { get; }
        ITemplateDefinition TemplateDefinition { get; }

        Dictionary<ITemplateDefinition, ISortersAndFilters> ExternalSortersAndFilters { get; set; }

        string SearchValue { get; set; }
        void ExecuteSearch();
    }
}
