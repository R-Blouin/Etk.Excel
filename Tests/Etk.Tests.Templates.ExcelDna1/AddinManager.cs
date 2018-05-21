using Etk.Excel;
using Etk.Tests.Templates.ExcelDna1.BasicExcelComTests;
using Etk.Tests.Templates.ExcelDna1.Dashboard;
using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Etk.Tests.Templates.ExcelDna1
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
            if(currentWorkbook != null)
                currentWorkbook.BeforeClose += (ref bool cancel) => currentWorkbook.Saved = true;

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);

            // Create, render and activate the dashboard view
            DashboardSheet.CreateAndActivateDashBoard();

            Marshal.ReleaseComObject(currentWorkbook);

            BasicExcelComTestsManager basicExcelComTests = new BasicExcelComTestsManager();
            basicExcelComTests.Execute();
        }

        public void AutoClose()
        { }
    }
}
