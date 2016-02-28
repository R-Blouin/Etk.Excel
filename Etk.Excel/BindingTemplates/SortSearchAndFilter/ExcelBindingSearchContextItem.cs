namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    using System.Collections.Generic;
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
    using Etk.BindingTemplates.Views;
    using Etk.Excel.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;
    using Etk.Excel.Application;
    using Etk.BindingTemplates.Definitions.Templates;

    class ExcelBindingSearchContextItem : BindingSearchContextItem
    {
        public ExcelBindingSearchContextItem(ITemplateView view, BindingSearchDefinition definition)
                                            : base(view, definition)
        { }

        override protected void ExecuteSearch(ITemplateView view)
        {
            using (FreezeExcel freezeExcel = new FreezeExcel())
            {
                if (view == null || ! (view is ExcelTemplateView)
                   || ((ExcelTemplateView)view).Renderer == null 
                   || ((ExcelTemplateView)view).Renderer.BodyPartRenderer == null
                   || ((ExcelTemplateView)view).Renderer.BodyPartRenderer.RenderedArea == null)
                    return;

                ExcelTemplateView excelView = (ExcelTemplateView) view;
                
                List<KeyValuePair<Range, bool>> toShowOrHide = new  List<KeyValuePair<Range, bool>>();

                Range firstRange = excelView.SheetDestination.Cells[excelView.Renderer.BodyPartRenderer.RenderedArea.YPos, excelView.Renderer.BodyPartRenderer.RenderedArea.XPos];
                Range lastRange = excelView.SheetDestination.Cells[excelView.Renderer.BodyPartRenderer.RenderedArea.YPos + excelView.Renderer.BodyPartRenderer.RenderedArea.Height - 1, excelView.Renderer.BodyPartRenderer.RenderedArea.XPos + excelView.Renderer.BodyPartRenderer.RenderedArea.Width - 1];
                Range renderedRange = excelView.SheetDestination.Range[firstRange, lastRange];
                Range rowsOrColumns = view.TemplateDefinition.Orientation == Orientation.Horizontal ? renderedRange.Columns : renderedRange.Cells.Rows;
                if (string.IsNullOrEmpty(excelView.SearchValue))
                {
                    foreach (Range rowOrColumn in rowsOrColumns)
                        toShowOrHide.Add(new KeyValuePair<Range, bool>(rowOrColumn, false));
                }
                else
                {
                    string searchValueUpper = excelView.SearchValue.ToUpper();
                    foreach (Range rowOrColumn in rowsOrColumns)
                    {
                        bool toHide = true;
                        foreach(Range cell in rowOrColumn.Cells)
                        {
                            string cellText = cell.Text;
                            if (!string.IsNullOrEmpty(cellText) && cellText.ToUpper().Contains(searchValueUpper))
                            {
                                toHide = false;
                                break;
                            }
                        }
                        toShowOrHide.Add(new KeyValuePair<Range, bool>(rowOrColumn, toHide));
                    }
                }
                HideShowRanges(excelView, toShowOrHide, false);
            }
        }

        private void HideShowRanges(ExcelTemplateView view, List<KeyValuePair<Range, bool>> toShowOrHide, bool hide)
        {
            foreach (KeyValuePair<Range, bool> showOrHide in toShowOrHide)
            {
                Range cells;
                if (view.TemplateDefinition.Orientation == Orientation.Horizontal)
                    cells = view.SheetDestination.Columns[showOrHide.Key.Column];
                else
                    cells = view.SheetDestination.Rows[showOrHide.Key.Row];
                cells.Hidden = showOrHide.Value;
                cells = null;
            }
        }
    }
}
