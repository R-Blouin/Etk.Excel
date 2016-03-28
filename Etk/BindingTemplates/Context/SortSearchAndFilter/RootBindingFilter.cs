using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context.SortSearchAndFilter
{
    public class RootBindingFilter : IFilterDefinition
    {
        #region attributes and properties
        private TemplateView view;
        private readonly BindingFilterDefinition bindingFilterDefinition;

        public ITemplateDefinition TemplateDefinition
        { get { return bindingFilterDefinition.FilterOwner.Parent; } }

        public string FilterExpression
        { get { return bindingFilterDefinition.GetFilterExpression(FilterValue); } }

        public string FilterValue
        { get; private set; }

        public IBindingDefinition DefinitionToFilter
        { get { return bindingFilterDefinition.DefinitionToFilter; } }
        #endregion

        #region .ctors and factories
        private RootBindingFilter(ITemplateView view, BindingFilterDefinition bindingFilterDefinition, string filterValue)
        {
            this.view = (TemplateView)view;
            this.bindingFilterDefinition = bindingFilterDefinition;
            FilterValue = filterValue;
        }

        public static List<IFilterDefinition> CreateInstances(TemplateView view, object dataSource)
        {
            Dictionary<BindingFilterDefinition, string> filterDefinitionByElement;
            if (view.FilterValueByFilterDefinitionByElement.TryGetValue(dataSource, out filterDefinitionByElement))
            {
                IEnumerable<KeyValuePair<BindingFilterDefinition, string>> activeFilters = filterDefinitionByElement.Where(f => !string.IsNullOrEmpty(f.Value));
                return activeFilters.Select(af => (IFilterDefinition) new RootBindingFilter(view, af.Key, af.Value)).ToList();
            }
            return null;
        }
        #endregion
    }
}
