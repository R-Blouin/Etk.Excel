namespace Etk.Tests.Templates.ExcelDna1.Dashboard
{
    using EEtk.Tests.Templates.ExcelDna1.Dashboard;
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using ExcelDna.Integration.CustomUI;
    using System.IO;
    using System.Reflection;

    class DashboardSheet
    {
        private IExcelTemplateView view;
        private CustomTaskPane taskPane;

        #region .ctors and factories
        private DashboardSheet()
        {}

        public static void CreateAndActivateDashBoard()
        {
            DashboardSheet dashboardSheet = new DashboardSheet();

            dashboardSheet.DeclareDecorators();
            dashboardSheet.CreateAndRender();

            dashboardSheet.CreateDashboardTaskPane();

            dashboardSheet.view.ViewSheetIsActivated += dashboardSheet.OnSheetIsActivated;
            dashboardSheet.view.ViewSheetIsDeactivated += dashboardSheet.OnSheetIsDeactivated;

            // Insure the dashboard sheet is activated
            dashboardSheet.view.ViewSheet.Activate();
        }
        #endregion

        #region 
        private void OnSheetIsActivated(IExcelTemplateView notUsedParameter)
        {
            taskPane.Visible = true;
        }

        private void OnSheetIsDeactivated(IExcelTemplateView notUsedParameter)
        {
            taskPane.Visible = false;
        }

        /// <summary> Declare the decorator used in the dashboard</summary>
        private void DeclareDecorators()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Tests.Templates.ExcelDna1.Dashboard.DashboardDecoratorDefinitions.xml")))
            {
                ETKExcel.TemplateManager.RegisterDecoratorsFromXml(textReader.ReadToEnd());
            }
        }

        private void CreateDashboardTaskPane()
        {
            // Create and display a taskpane => to see the interaction between data in Wpf UI and ETK templates
            taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(DashboardPanel), "Dashboard Panel");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
        }

        /// <summary> Create and render the dashboard view</summary>
        private void CreateAndRender()
        {
            view = ETKExcel.TemplateManager.AddView("Dashboard Templates", "Main", "Dashboard", "B2");
            // Inject the data source
            ExcelTestsManager testsManager = new ExcelTestsManager();
            view.SetDataSource(testsManager);
            // RenderView the dashboard
            ETKExcel.TemplateManager.Render(view);
        }
        #endregion
    }
}
