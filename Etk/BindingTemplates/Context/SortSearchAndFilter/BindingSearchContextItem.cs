using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Views;

namespace Etk.BindingTemplates.Context.SortSearchAndFilter
{
    public abstract class BindingSearchContextItem : BindingContextItem
    {
        #region attributes and properties
        private readonly TemplateView view;
        private readonly BindingSearchDefinition definition;
        #endregion

        #region .ctors
        protected BindingSearchContextItem(ITemplateView view, BindingSearchDefinition definition)
                                          : base(null, null)                
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
            ExecuteSearch(view);
            return true;
        }
        #endregion

        #region protected methods
        protected abstract void ExecuteSearch(ITemplateView view);
        #endregion
    }
}