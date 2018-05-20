using EEtk.Tests.Templates.ExcelDna1.Dashboard;
using Etk.Excel;
using Etk.Excel.BindingTemplates.Views;
using ExcelDna.Integration.CustomUI;
using System.IO;
using System.Reflection;
using Etk.Excel.Application;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System;
using System.Diagnostics;

namespace Etk.Tests.Templates.ExcelDna1.Dashboard
{
    class DashboardSheet
    {
        private IExcelTemplateView view;
        private CustomTaskPane taskPane;

        #region .ctors and factories
        private DashboardSheet()
        { }

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
        private void OnSheetIsActivated()
        {
            using (FreezeExcel freeExcel = new FreezeExcel())
            {
                taskPane.Visible = true;
            }
        }

        private void OnSheetIsDeactivated()
        {
            using (FreezeExcel freeExcel = new FreezeExcel())
            {
                taskPane.Visible = false;
            }
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
            view = ETKExcel.TemplateManager.AddView("Dashboard Templates", "Main", "Dashboard", "B1");


            //Test1();
            //Test2();
            Test3();

            // Inject the data source
            ExcelTestsManager testsManager = new ExcelTestsManager();
            view.SetDataSource(testsManager);
            // RenderView the dashboard
            ETKExcel.TemplateManager.Render(view);
        }
        #endregion


        void ReleaseComObject(object obj)
        {
            int refCpt = Marshal.ReleaseComObject(obj);
            Trace.WriteLine($"Marshal cpt: {refCpt}");
            if (refCpt < 0)
                Trace.WriteLine("Aie !!! ReleaseComObject");
        }

        void Test1()
        {
            ExcelInterop.Range range = view.ViewSheet.Range["A1"];
            Test1_1(range);
            ReleaseComObject(range);
        }

        void Test1_1(ExcelInterop.Range range)
        {
            ExcelInterop.Borders borders = range.Borders;

            ////borders.Color = color;
            ReleaseComObject(borders);
            borders = null;
        }

        public void Test2()
        {
            ExcelInterop.Workbook workbook = null;
            ExcelInterop.Sheets sheets = null;
            ExcelInterop.Worksheet lastSheet = null;
            ExcelInterop.Worksheet firstSheet = null;
            try
            {
                //viewsOwnerSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, DestinationSheetName);

                // Create the destination sheet
                workbook = ETKExcel.ExcelApplication.Application.ActiveWorkbook;
                //sheets = workbook.Sheets;
                //firstSheet = workbook.Sheets[1];
                //lastSheet = workbook.Sheets[sheets.Count];
            }
            finally
            {
                if (firstSheet != null)
                    ReleaseComObject(firstSheet);
                if (lastSheet != null)
                    ReleaseComObject(lastSheet);
                if (sheets != null)
                    ReleaseComObject(sheets);
                if (workbook != null)
                    ReleaseComObject(workbook);
            }
        }

        public void Test3()
        {
            ExcelInterop.Range range = view.ViewSheet.Range["A1"];
            Test3_3(range, 1);
            Test3_3(range, 2);
            ReleaseComObject(range);
        }

        public void Test3_3(ExcelInterop.Range targetedRange, int numberOfColumns)
        {
            ExcelInterop.Range workingRange;
            if (numberOfColumns < 0)
                workingRange = targetedRange.Offset[Type.Missing, numberOfColumns];
            else
                workingRange = targetedRange.Offset[Type.Missing, 1];

            workingRange = workingRange.Resize[Type.Missing, Math.Abs(numberOfColumns)];
            workingRange = workingRange.Resize[Type.Missing, Math.Abs(numberOfColumns)];
            workingRange = workingRange.Resize[Type.Missing, Math.Abs(numberOfColumns)];

            ExcelInterop.Range columns = workingRange.EntireColumn;
            columns.Hidden = !(bool)columns.Hidden;

            ReleaseComObject(columns);
            ReleaseComObject(workingRange);
        }
    }
}
