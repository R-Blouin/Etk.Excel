using Etk.Excel.BindingTemplates.Views;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Addin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class EtkExcelAddin
    {
        /// <summary> Create a View.</summary>
        public IExcelTemplateView AddView(ExcelInterop.Worksheet sheetContainer, string templateName, ExcelInterop.Worksheet sheetDestination, ExcelInterop.Range destinationRange, ExcelInterop.Range clearingCell = null)
        {
            return ETKExcel.TemplateManager.AddView(sheetContainer, templateName, sheetDestination, destinationRange, clearingCell);
        }

        public IExcelTemplateView AddViewFromNames(string sheetTemplatePath, string templateName, string sheetDestinationName, string destinationRange, string clearingCell = null)
        {
            return ETKExcel.TemplateManager.AddView(sheetTemplatePath, templateName, sheetDestinationName, destinationRange, clearingCell);
        }

        public void SetDataSource(IExcelTemplateView view, object dataSource)
        {
            view.SetDataSource(dataSource);
        }
        

        /// <summary> Remove and dispose a list of View.</summary>
        public void RemoveViews(IEnumerable<IExcelTemplateView> views)
        {
            ETKExcel.TemplateManager.RemoveViews(views);
        }

        /// <summary> Remove a View.</summary>
        public void RemoveView(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.RemoveView(view);
        }

        /// <summary> Get all the views owned by a given sheet.</summary>
        public IExcelTemplateView[] GetSheetViews(ExcelInterop.Worksheet sheet)
        {
            IEnumerable<IExcelTemplateView>  views = ETKExcel.TemplateManager.GetSheetViews(sheet);
            return views?.ToArray();
        }

        /// <summary> Get all the views owned by a the current active sheet.</summary>
        public IExcelTemplateView[] GetActiveSheetViews()
        {
            IEnumerable<IExcelTemplateView> views = ETKExcel.TemplateManager.GetActiveSheetViews();
            return views?.ToArray();
        }

        /// <summary> Rerender (description, style and data) the View given as parameter.</summary>
        public void RenderView(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.Render(view);
        }

        /// <summary> Rerender (description, style and data) all the views given as parameters.</summary>
        public void RenderViews(IEnumerable<IExcelTemplateView> views)
        {
            ETKExcel.TemplateManager.Render(views);
        }

        /// <summary> Rerender only the data of the View given as parameter 
        public void RenderViewDataOnly(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.Render(view);
        }

        /// <summary> Rerender only the data of all the views given as parameters 
        public void RenderViewsDataOnly(IEnumerable<IExcelTemplateView> views)
        {
            ETKExcel.TemplateManager.Render(views);
        }

        /// <summary> Clear the previously rendering View
        /// <param name="View">The View to clear.</param>
        public void ClearView(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.ClearView(view);
        }

        /// <summary> Clear the previously rendering views
        /// <param name="views">The views to clear.</param>
        public void ClearViews(IEnumerable<IExcelTemplateView> views)
        {
            ETKExcel.TemplateManager.ClearViews(views);
        }


        ///// <summary> Register decorator definitions from a xml
        ///// <param name="xmLDefinition">The xml data containing the decorator definitions</param>
        //public void RegisterDecoratorsFromXml(string xmLDefinition)
        //{
        //}

        ///// <summary> Register decorator definitions
        ///// <param name="rangeDecorator">the rangeDecorator to register</param>
        //public void RegisterDecorator(ExcelRangeDecorator rangeDecorator)
        //{
        //}

        ///// <summary> Register Event callback definitions  from a xml
        ///// <param name="xmLDefinition">The xml data containing the callback definitions</param>
        //public void RegisterEventCallbacksFromXml(string xmLDefinition)
        //{
        //}

        ///// <summary> Register Event callback definitions 
        ///// <param name="callback">The callback to register</param>
        //public void RegisterEventCallback(EventCallback callback)
        //{
        //}
    }
}
