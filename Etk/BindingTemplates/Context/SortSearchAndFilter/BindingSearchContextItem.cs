namespace Etk.BindingTemplates.Context.SortSearchAndFilter
{
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
    using Etk.BindingTemplates.Views;

    public abstract class BindingSearchContextItem : BindingContextItem
    {
        #region attributes and properties
        private TemplateView view;
        private BindingSearchDefinition definition;
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
        override public object ResolveBinding()
        {
            return string.IsNullOrEmpty(view.SearchValue) ? definition.Watermark : view.SearchValue;
        }

        override public bool UpdateDataSource(object data, out object retValue)
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
        abstract protected void ExecuteSearch(ITemplateView view);
        #endregion
    }
}