using System.Runtime.InteropServices;

namespace Etk.Demo.Shops.UI.Excel
{
    using Etk.Demo.Shops.UI.Excel.Sheets;
    using Etk.Excel;
    using ExcelDna.Integration;
    using Excel = Microsoft.Office.Interop.Excel;

    class AddinManager : IExcelAddIn
    {
        public void AutoOpen()
        {
            Excel.Application excelApplication = ExcelDnaUtil.Application as Excel.Application;

            // To avoid the Excel 'Save message' on Exit
            Excel.Workbook currentWorkbook = excelApplication.ActiveWorkbook;
            currentWorkbook.BeforeClose += (ref bool cancel) =>
            {
                currentWorkbook.Saved = true;
            };

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(excelApplication);

            // Create, render shoop view
            SheetShopsRef.Instance.RenderViews();
        }

        public void AutoClose()
        { }
    }
}
