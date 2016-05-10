namespace Etk.Tests.Templates.ExcelDna1
{
    using System.Collections.Generic;
    using Etk.Excel;
    using Etk.Tests.Templates.ExcelDna1.Tests;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter;
    using Etk.Tools.Patterns;
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    class ExcelTestsManager
    {
        #region attributes and properties 
        private List<IExcelTests> tests;

        public IEnumerable<IExcelTests> Tests
        { get { return tests; } }
        #endregion

        #region .Ctors
        public ExcelTestsManager()
        {
            tests = new List<IExcelTests>();
            tests.Add(new BasicVerticalNoHeaderAndFooterTests());
            tests.Add(new BasicVerticalMonoHeaderAndFooterTests());
        }
        #endregion

        #region public methods
        /// <summary>
        /// Tests execution
        /// </summary>
        public void Execute()
        {
            ETKExcel.ExcelApplication.PostAsynchronousAction(() => 
                    {
                        foreach (IExcelTests test in Tests)
                            test.Execute();

                        ExcelInterop.Worksheet dashBoardSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, "Dashboard");
                        if (dashBoardSheet != null)
                            dashBoardSheet.Activate();
                    });
        }
        #endregion
    }
}
