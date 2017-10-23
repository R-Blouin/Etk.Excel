namespace Etk.Tests.Templates.ExcelDna1
{
    using Dashboard;
    using Etk.Excel;
    using ExcelDna.Integration;
    using Excel = Microsoft.Office.Interop.Excel;

    class AddinManager : IExcelAddIn
    {
        public Excel.Application ExcelApplication
        { get; private set; }

        public void AutoOpen()
        {
            ExcelApplication = ExcelDnaUtil.Application as Excel.Application;

            // To avoid the Excel 'Save message' on Exit
            Excel.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;
            if(currentWorkbook != null)
                currentWorkbook.BeforeClose += (ref bool cancel) => currentWorkbook.Saved = true;

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);

            // Create, render and activate the dashboard view
            DashboardSheet.CreateAndActivateDashBoard();
        }

        public void AutoClose()
        { }
    }
}
