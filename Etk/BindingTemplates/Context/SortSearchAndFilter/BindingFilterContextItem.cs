using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context.SortSearchAndFilter
{
    public abstract class BindingFilterContextItem : BindingContextItem, IFilterDefinition
    {
        #region attributes and properties
        private readonly TemplateView view;
        private readonly IBindingContextElement bindingContextElement;
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

        #region .ctors
        protected BindingFilterContextItem(ITemplateView view, BindingFilterDefinition bindingFilterDefinition, IBindingContextElement bindingContextElement)
                                          : base(bindingContextElement, null)                
        {
            this.view = (TemplateView)view;
            this.bindingFilterDefinition = bindingFilterDefinition;
            this.bindingContextElement = bindingContextElement;

            Dictionary<BindingFilterDefinition, string> filterDefinitionByElement;
            if (! ((TemplateView)view).FilterValueByFilterDefinitionByElement.TryGetValue(bindingContextElement.DataSource, out filterDefinitionByElement))
            {
                filterDefinitionByElement = new Dictionary<BindingFilterDefinition, string>();
                ((TemplateView)view).FilterValueByFilterDefinitionByElement[bindingContextElement.DataSource] = filterDefinitionByElement;           
            }

            string filterValue;
            filterDefinitionByElement.TryGetValue(bindingFilterDefinition, out filterValue);

            FilterValue = filterValue;
        }
        #endregion

        #region public methods
        public override object ResolveBinding()
        {
            return string.IsNullOrEmpty(FilterValue) ? bindingFilterDefinition.Watermark : FilterValue;
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            Dictionary<BindingFilterDefinition, string> filterDefinitionByElement;
            if (!view.FilterValueByFilterDefinitionByElement.TryGetValue(bindingContextElement.DataSource, out filterDefinitionByElement))
                throw new Exception("Update Filter set in Template failed. Cannot retrieve the saved filtering expression...");

            if (data != null)
            {
                filterDefinitionByElement[bindingFilterDefinition] = FilterValue = data.ToString();
                retValue = FilterValue;
            }
            else
            {
                filterDefinitionByElement[bindingFilterDefinition] = FilterValue = null;
                retValue = bindingFilterDefinition.Watermark;
            }
            ExecuteFilter(view);
            return true;
        }
        #endregion

        #region protected methods
        protected abstract void ExecuteFilter(ITemplateView view);
        #endregion
    }
}