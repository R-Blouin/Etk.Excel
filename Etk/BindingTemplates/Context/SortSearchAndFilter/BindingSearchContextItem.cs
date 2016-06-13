using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Views;

namespace Etk.BindingTemplates.Context.SortSearchAndFilter
{
    public abstract class BindingSearchContextItem : BindingContextItem
    {
        #region attributes and properties
        protected readonly TemplateView view;
        protected readonly BindingSearchDefinition definition;
        #endregion

        #region .ctors
        protected BindingSearchContextItem(ITemplateView view, BindingSearchDefinition definition, IBindingContextElement parent)
                                          : base(parent, null)                
        {
            this.view = (TemplateView)view;
            this.definition = definition;
        }
        #endregion

        #region public methods
        public override object ResolveBinding()
        {
            return string.IsNullOrEmpty(view.SearchValue) ? definition.Watermark : view.SearchValue;
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            if (data != null)
            {
                view.SearchValue = data.ToString();
                retValue = view.SearchValue;
            }
            else
            {
                view.SearchValue = null;
                retValue = definition.Watermark;
            }
            return true;
        }

        protected void ExecuteSearch()
        {
            ((TemplateView)view).ExecuteSearch();
        }
        #endregion
    }
}