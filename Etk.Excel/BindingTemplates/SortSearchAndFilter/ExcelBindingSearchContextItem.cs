using System.Collections.Generic;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    class ExcelBindingSearchContextItem : BindingSearchContextItem
    {
        public ExcelBindingSearchContextItem(ITemplateView view, BindingSearchDefinition definition)
                                            : base(view, definition)
        { }

        protected override void ExecuteSearch(ITemplateView view)
        {
            using (FreezeExcel freezeExcel = new FreezeExcel())
            {
                if (view == null || ! (view is ExcelTemplateView)
                   || ((ExcelTemplateView)view).Renderer == null 
                   || ((ExcelTemplateView)view).Renderer.BodyPartRenderer == null
                   || ((ExcelTemplateView)view).Renderer.BodyPartRenderer.RenderedArea == null)
                    return;

                ExcelTemplateView excelView = (ExcelTemplateView) view;

                List<KeyValuePair<ExcelInterop.Range, bool>> toShowOrHide = new List<KeyValuePair<ExcelInterop.Range, bool>>();

                ExcelInterop.Range firstRange = excelView.SheetDestination.Cells[excelView.Renderer.BodyPartRenderer.RenderedArea.YPos, excelView.Renderer.BodyPartRenderer.RenderedArea.XPos];
                ExcelInterop.Range lastRange = excelView.SheetDestination.Cells[excelView.Renderer.BodyPartRenderer.RenderedArea.YPos + excelView.Renderer.BodyPartRenderer.RenderedArea.Height - 1, excelView.Renderer.BodyPartRenderer.RenderedArea.XPos + excelView.Renderer.BodyPartRenderer.RenderedArea.Width - 1];
                ExcelInterop.Range renderedRange = excelView.SheetDestination.Range[firstRange, lastRange];
                ExcelInterop.Range rowsOrColumns = view.TemplateDefinition.Orientation == Orientation.Horizontal ? renderedRange.Columns : renderedRange.Cells.Rows;
                if (string.IsNullOrEmpty(excelView.SearchValue))
                {
                    foreach (ExcelInterop.Range rowOrColumn in rowsOrColumns)
                        toShowOrHide.Add(new KeyValuePair<ExcelInterop.Range, bool>(rowOrColumn, false));
                }
                else
                {
                    string searchValueUpper = excelView.SearchValue.ToUpper();
                    foreach (ExcelInterop.Range rowOrColumn in rowsOrColumns)
                    {
                        bool toHide = true;
                        foreach (ExcelInterop.Range cell in rowOrColumn.Cells)
                        {
                            string cellText = cell.Text;
                            if (!string.IsNullOrEmpty(cellText) && cellText.ToUpper().Contains(searchValueUpper))
                            {
                                toHide = false;
                                break;
                            }
                        }
                        toShowOrHide.Add(new KeyValuePair<ExcelInterop.Range, bool>(rowOrColumn, toHide));
                    }
                }
                HideShowRanges(excelView, toShowOrHide, false);
            }
        }

        private void HideShowRanges(ExcelTemplateView view, List<KeyValuePair<ExcelInterop.Range, bool>> toShowOrHide, bool hide)
        {
            foreach (KeyValuePair<ExcelInterop.Range, bool> showOrHide in toShowOrHide)
            {
                ExcelInterop.Range cells;
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
