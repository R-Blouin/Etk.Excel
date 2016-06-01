namespace Etk.Tests.Templates.ExcelDna1
{
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using ExcelDna.Integration;
    using Tests;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Reflection;
    using System.IO;

    class AddinManager : IExcelAddIn
    {
        public Excel.Application ExcelApplication
        { get; private set; }

        public void AutoOpen()
        {
            ExcelApplication = ExcelDnaUtil.Application as Excel.Application;

            // To avoid the Excel 'Save message' on Exit
            Excel.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;
            currentWorkbook.BeforeClose += (ref bool cancel) => currentWorkbook.Saved = true;

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);

            Assembly assembly = Assembly.GetExecutingAssembly();
            // Declare the decorator used in the dashboad
            using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Tests.Templates.ExcelDna1.DashboardDecoratorDefinitions.xml")))
            {
                ETKExcel.TemplateManager.RegisterDecoratorsFromXml(textReader.ReadToEnd());
            }

            // Create the dashboard view
            IExcelTemplateView view = ETKExcel.TemplateManager.AddView("Dashboard Templates", "Main", "Dashboard", "B2");

            // Create a inject the data source
            ExcelTestsManager testsManager = new ExcelTestsManager();
            view.SetDataSource(testsManager);
            
            // Render the dashboard
            ETKExcel.TemplateManager.Render(view);

            // Activate the dashboard sheet
            Excel.Worksheet dashBoardSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, "Dashboard");
            if (dashBoardSheet != null)
                ((Excel._Worksheet) dashBoardSheet).Activate();
        }

        public void AutoClose()
        { }
    }
}
