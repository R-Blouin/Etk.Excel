using System.IO;
using Etk.Demos.Data.Shops;
using Etk.Excel;
using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.BindingTemplates.Views;

namespace Etk.Demo.ExcelDna1
{
    class AddinManager : IExcelAddIn
    {
        public ExcelInterop.Application ExcelApplication
        { get; private set; }

        public void AutoOpen()
        {
            ExcelApplication = ExcelDnaUtil.Application as ExcelInterop.Application;

            // To avoid the Excel 'Save message' on Exit
            ExcelInterop.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;
            currentWorkbook.BeforeClose += (ref bool cancel) => currentWorkbook.Saved = true;

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);

            ExcelInterop.Worksheet worksheetDestination = ETKExcel.ExcelApplication.GetWorkSheetFromName(currentWorkbook, "Results");

            string addinPath = (string)XlCall.Excel(XlCall.xlGetName);
            string rootPath = Path.GetDirectoryName(addinPath);
            ExcelInterop.Workbooks workbooks = ETKExcel.ExcelApplication.Application.Workbooks;
            ExcelInterop.Workbook workbookTemplateContainer = workbooks.Open(Path.Combine(rootPath, "Etk.Demo.Templates.xlsx"), true, true);

            ExcelInterop.Worksheet worksheetTemplateSource = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbookTemplateContainer, "Templates");
            ExcelInterop.Range viewFirstOutputRange = worksheetDestination.Range["B2"];
            IExcelTemplateView mainview  = ETKExcel.TemplateManager.AddView(worksheetTemplateSource, "Main", worksheetDestination, viewFirstOutputRange);
            mainview.SetDataSource(ShopManager.Shops);
            mainview.Render();
        }

        public void AutoClose()
        { }
    }
}
