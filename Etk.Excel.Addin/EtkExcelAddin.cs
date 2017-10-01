using Etk.Excel.BindingTemplates.Views;
using System;
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
            try
            {
                return ETKExcel.TemplateManager.AddView(sheetContainer, templateName, sheetDestination, destinationRange, clearingCell);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Add View' failed", ex);
                return null;
            }
        }

        public EtkView AddViewFromNames(string sheetTemplatePath, string templateName, string sheetDestinationName, string destinationRange, string clearingCell = null)
        {
            try
            {
                IExcelTemplateView view = ETKExcel.TemplateManager.AddView(sheetTemplatePath, templateName, sheetDestinationName, destinationRange, clearingCell);
                return EtkView.CreateInstance(view);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Add View from names' failed", ex);
                return null;
            }
        }

        public void SetDataSource(EtkView view, object dataSource)
        {
            try
            {
                if (view != null)
                    view.DataSource = dataSource;
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Set Data Source' failed", ex);
            }
        }
        

        /// <summary> Remove and dispose a list of View.</summary>
        public void RemoveViews(IEnumerable<EtkView> views)
        {
            try
            {
                if (views != null)
                    ETKExcel.TemplateManager.RemoveViews(views.Select(v => v.ExcelView));
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Remove Views' failed", ex);
            }
        }

        /// <summary> Remove a View.</summary>
        public void RemoveView(EtkView view)
        {
            try
            {
                ETKExcel.TemplateManager.RemoveView(view?.ExcelView);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Remove View' failed", ex);
            }
        }

        /// <summary> Get all the views owned by a given sheet.</summary>
        public EtkView[] GetSheetViews(ExcelInterop.Worksheet sheet)
        {
            IEnumerable<IExcelTemplateView> views = ETKExcel.TemplateManager.GetSheetViews(sheet);
            if (views == null)
                return null;
            return views.Select(v => EtkView.CreateInstance(v)).ToArray();
        }

        /// <summary> Get all the views owned by a the current active sheet.</summary>
        public EtkView[] GetActiveSheetViews()
        {
            IEnumerable<IExcelTemplateView> views = ETKExcel.TemplateManager.GetActiveSheetViews();
            if (views == null)
                return null;
            return views.Select(v => EtkView.CreateInstance(v)).ToArray();
        }

        /// <summary> Rerender (description, style and data) the View given as parameter.</summary>
        public void RenderView(EtkView view)
        {
            try
            {
                ETKExcel.TemplateManager.Render(view?.ExcelView);
            }
            catch(Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Render View' failed", ex);
            }
        }

        /// <summary> Rerender (description, style and data) all the views given as parameters.</summary>
        public void RenderViews(IEnumerable<EtkView> views)
        {
            try
            {
                if (views != null)
                    ETKExcel.TemplateManager.Render(views.Select(v => v.ExcelView));
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Render Views' failed", ex);
            }
        }

        /// <summary> Rerender only the data of the View given as parameter 
        public void RenderViewDataOnly(EtkView view)
        {
            try
            {
                ETKExcel.TemplateManager.RenderDataOnly(view?.ExcelView);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Render View Data Only' failed", ex);
            }
        }

        /// <summary> Rerender only the data of all the views given as parameters 
        public void RenderViewsDataOnly(IEnumerable<EtkView> views)
        {
            try
            {
                if (views != null)
                    ETKExcel.TemplateManager.RenderDataOnly(views.Select(v => v.ExcelView));
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Render Views Data Only' failed", ex);
            }
        }

        /// <summary> Clear the previously rendering View
        /// <param name="View">The View to clear.</param>
        public void ClearView(EtkView view)
        {
            try
            {
                ETKExcel.TemplateManager.ClearView(view?.ExcelView);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Clear View' failed", ex);
            }
        }

        /// <summary> Clear the previously rendering views
        /// <param name="views">The views to clear.</param>
        public void ClearViews(IEnumerable<EtkView> views)
        {
            try
            {
                if (views != null)
                    ETKExcel.TemplateManager.ClearViews(views.Select(v => v.ExcelView));
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, "'Clear Views' failed", ex);
            }
        }

        public void ClearRange(ExcelInterop.Range from, ExcelInterop.Range to = null, ExcelInterop.Range with = null)
        {
            ETKExcel.ExcelApplication.ClearRange(from, to, with);
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
