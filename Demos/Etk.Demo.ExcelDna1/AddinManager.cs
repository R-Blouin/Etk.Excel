using Etk.Demos.Data.Shops;
using Etk.Excel;
using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.BindingTemplates.Views;
using System.Runtime.InteropServices;
using Etk.Demos.Data.Shares;

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

            ShowShops();
            ShowShares();
        }

        private void ShowShops()
        {
            ExcelInterop.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;

            //string addinPath = (string)XlCall.Excel(XlCall.xlGetName);
            //string rootPath = Path.GetDirectoryName(addinPath);
            //ExcelInterop.Workbooks workbooks = ETKExcel.ExcelApplication.Application.Workbooks;
            //ExcelInterop.Workbook workbookTemplateContainer = workbooks.Open(Path.Combine(rootPath, "Etk.Demo.Templates.xlsx"), true, true);

            ExcelInterop.Worksheet worksheetDestination = ETKExcel.ExcelApplication.GetWorkSheetFromName(currentWorkbook, "Shops");
            //ExcelInterop.Worksheet worksheetTemplateSource = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbookTemplateContainer, "Templates");
            ExcelInterop.Range viewFirstOutputRange = worksheetDestination.Range["B2"];
            IExcelTemplateView mainview = ETKExcel.TemplateManager.AddView("TemplatesShops", "Main", "Shops", "B2");

            ShopsManager shopsManager = new ShopsManager();
            mainview.SetDataSource(shopsManager.Shops);
            mainview.Render();


            Marshal.ReleaseComObject(currentWorkbook);
            Marshal.ReleaseComObject(worksheetDestination);
        }

        private void ShowShares()
        {
            new BasketManager();
            ExcelInterop.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;

            //string addinPath = (string)XlCall.Excel(XlCall.xlGetName);
            //string rootPath = Path.GetDirectoryName(addinPath);
            //ExcelInterop.Workbooks workbooks = ETKExcel.ExcelApplication.Application.Workbooks;
            //ExcelInterop.Workbook workbookTemplateContainer = workbooks.Open(Path.Combine(rootPath, "Etk.Demo.Templates.xlsx"), true, true);

            ExcelInterop.Worksheet worksheetDestination = ETKExcel.ExcelApplication.GetWorkSheetFromName(currentWorkbook, "Shares");
            //ExcelInterop.Worksheet worksheetTemplateSource = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbookTemplateContainer, "Templates");
            ExcelInterop.Range viewFirstOutputRange = worksheetDestination.Range["B2"];
            IExcelTemplateView mainview = ETKExcel.TemplateManager.AddView("TemplatesShares", "Main", "Shares", "B2");

            BasketManager basketManager = new BasketManager();
            mainview.SetDataSource(basketManager);
            mainview.Render();
            mainview.ViewSheetIsActivated += () => basketManager.StartChanging();
            mainview.ViewSheetIsDeactivated += () => basketManager.StopChanging();


            Marshal.ReleaseComObject(currentWorkbook);
            Marshal.ReleaseComObject(worksheetDestination);
        }

        public void AutoClose()
        { }
    }
}
