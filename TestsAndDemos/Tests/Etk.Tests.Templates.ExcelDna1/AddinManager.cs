namespace Etk.Tests.Templates.ExcelDna1
{
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using ExcelDna.Integration;
    using Tests;
    using Excel = Microsoft.Office.Interop.Excel;

    class AddinManager : IExcelAddIn
    {
        private BasicViewsTestsManager basicViewsTestsManager = null;

        static public Excel.Application ExcelApplication
        { get; private set; }

        public void AutoOpen()
        {
            ExcelApplication = ExcelDnaUtil.Application as Excel.Application;

            // To avoid the Excel 'Save message' on Exit
            Excel.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;
            currentWorkbook.BeforeClose += (ref bool cancel) => currentWorkbook.Saved = true;

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);

            ExcelTestsManager testsManager = new ExcelTestsManager();

            IExcelTemplateView view = ETKExcel.TemplateManager.AddView("Dashboard Templates", "Main", "Dashboard", "B2");
            view.SetDataSource(testsManager);
            ETKExcel.TemplateManager.Render(view);
            //ExcelTestsManager.Instance.Execute();
        }

        public void AutoClose()
        {}
    }
}
