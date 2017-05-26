using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Views;

namespace Etk.Excel.BindingTemplates
{
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    public interface IExcelTemplateManager
    {
        /// <summary> Create a View.</summary>
        /// <param name="sheetContainer">The sheet that contains the description of the template to use.</param>
        /// <param name="templateName">The name of the template to use.</param>
        /// <param name="sheetDestination">The sheet where to render the View.</param>
        /// <param name="destinationRange">The first concernedRange where to render the View.</param>
        /// <param name="clearingCell">Optional: the cell uses to clear the View</param>
        /// <returns>The newly created View.</returns>
        IExcelTemplateView AddView(ExcelInterop.Worksheet sheetContainer, string templateName, ExcelInterop.Worksheet sheetDestination, ExcelInterop.Range destinationRange, ExcelInterop.Range clearingCell = null);

        /// <summary> Create a View.</summary>
        /// <param name="sheetTemplateName">The name of the sheet that contains the description of the template to use.</param>
        /// <param name="templateName">The name of the template to use.</param>
        /// <param name="sheetDestinationName">The name of the sheet where to render the View.</param>
        /// <param name="destinationRange">The first concerned Range where to render the View.</param>
        /// <param name="clearingRange">Optional: the cell uses to clear the View</param>
        /// <returns>The newly created View.</returns>
        IExcelTemplateView AddView(string sheetTemplateName, string templateName, string sheetDestinationName, string destinationRange, string clearingCell = null);

        /// <summary> Remove and dispose a list of View.</summary>
        /// <param name="templates">The views to remove and dispose.</param>
        void RemoveViews(IEnumerable<IExcelTemplateView> views);

        /// <summary> Remove a View.</summary>
        /// <param name="template">The View to remove and dispose.</param>
        void RemoveView(IExcelTemplateView view);

        /// <summary> Get all the views owned by a given sheet.</summary>
        /// <param name="sheet">The source sheet.</param>
        /// <returns>A collection containing the views owned by the sheet or an empty collection if no views found in the sheet.</returns>
        IEnumerable<IExcelTemplateView> GetSheetViews(ExcelInterop.Worksheet sheet);

        /// <summary> Get all the views owned by a the current active sheet.</summary>
        /// <returns>A collection containing the views owned by the sheet or an empty collection if no views found in the sheet.</returns>
        IEnumerable<IExcelTemplateView> GetActiveSheetViews();

        /// <summary> Rerender (description, style and data) the View given as parameter.</summary>
        /// <param name="View">The views to refresh.</param>
        void Render(IExcelTemplateView view);
        
        /// <summary> Rerender (description, style and data) all the views given as parameters.</summary>
        /// <param name="views">The views to refresh.</param>
        void Render(IEnumerable<IExcelTemplateView> views);

        /// <summary> Rerender only the data of the View given as parameter 
        /// (This implies that no change was made in the structure of the data).</summary>
        /// <param name="View">The views to refresh.</param>
        void RenderDataOnly(IExcelTemplateView view);

        /// <summary> Rerender only the data of all the views given as parameters 
        /// (This implies that no change was made in the structure of the data).</summary>
        /// <param name="views">The views to refresh.</param>
        void RenderDataOnly(IEnumerable<IExcelTemplateView> views);

        /// <summary> Clear the previously rendering View
        /// <param name="View">The View to clear.</param>
        void ClearView(IExcelTemplateView view);
        
        /// <summary> Clear the previously rendering views
        /// <param name="views">The views to clear.</param>
        void ClearViews(IEnumerable<IExcelTemplateView> views);

         /// <summary> Register decorator definitions from a xml
        /// <param name="xmLDefinition">The xml data containing the decorator definitions</param>
        void RegisterDecoratorsFromXml(string xmLDefinition);

        /// <summary> Register decorator definitions
        /// <param name="rangeDecorator">the rangeDecorator to register</param>
        void RegisterDecorator(ExcelRangeDecorator rangeDecorator);

        /// <summary> Register Event callback definitions  from a xml
        /// <param name="xmLDefinition">The xml data containing the callback definitions</param>
        void RegisterEventCallbacksFromXml(string xmLDefinition);

        /// <summary> Register Event callback definitions 
        /// <param name="callback">The callback to register</param>
        void RegisterEventCallback(EventCallback callback);

        /// <summary>
        /// Returns template details from template identifed by name given in parameter.
        /// </summary>
        /// <param name="sheetName">Search template's name</param>
        /// <returns>Template details from template identifed by name given in parameter.</returns>
        IEnumerable<IExcelTemplateDetails> GetTemplateDetails(string sheetName);
    }
}
