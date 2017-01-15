using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Views;
using Etk.Excel.BindingTemplates.Views;

namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    class ExcelBindingFilterContextItem : BindingFilterContextItem 
    {
        public ExcelBindingFilterContextItem(ITemplateView view, BindingFilterDefinition bindingFilterDefinition, IBindingContextElement bindingContextElement)
                                            : base(view, bindingFilterDefinition, bindingContextElement)
        { }

        protected override void ExecuteFilter(ITemplateView view)
        {
            object dataSource = view.GetDataSource();
            ETKExcel.TemplateManager.ClearView(view as ExcelTemplateView);
            // We reinject the datasource to force the filtering
            ((TemplateView)view).CreateBindingContext(dataSource);
            // RenderView the view to see the filering application
            ETKExcel.TemplateManager.Render(view as ExcelTemplateView);
        }
    }
}
