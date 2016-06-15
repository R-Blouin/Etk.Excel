namespace Etk.Tests.Templates.ExcelDna1
{
    using System.Collections.Generic;
    using Etk.BindingTemplates.Views;
    using Etk.Excel;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Tests;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicEtkFeatures;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMultiHeaderAndFooter;
    using Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter;
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    interface IExcelTestsManager
    {
        void ExecuteTopic(IExcelTestTopic topic);
    }

    class ExcelTestsManager : ViewModelBase, IExcelTestsManager 
    {
        #region attributes and properties 
        private List<IExcelTestTopic> testTopics;

        public IEnumerable<IExcelTestTopic> TestTopics
        { get { return testTopics; } }

        private string status;
        public string Status
        {
            get { return status; }
            private set
            {
                status = value;
                OnPropertyChanged("Status");
            }
        }
        #endregion

        #region .Ctors
        public ExcelTestsManager()
        {
            testTopics = new List<IExcelTestTopic>();
         
            testTopics.Add(new BasicEtkFeaturesTests(this));
            testTopics.Add(new BasicVerticalNoHeaderAndFooterTests(this));
            testTopics.Add(new BasicVerticalMonoHeaderAndFooterTests(this));
            testTopics.Add(new BasicVerticalMultiHeaderAndFooterTests(this));
        }
        #endregion

        #region public methods
        public static void SetSearchValue(ITemplateView concernedView, IExcelTestTopic topic)
        {
            concernedView.SearchValue = concernedView.SearchValue == topic.Description ? null : topic.Description;
            concernedView.ExecuteSearch();
        }

        /// <summary>
        /// Execution of all the tests declared on the test topics declared on this class (property 'TestTopics').
        ///<br/>Invoke by double-clicking on the template button 'ExecuteTopic All TestTopics' on the template 'Main' declared on the sheet 'Dashboard Templates'.
        /// </summary>
        public void Execute()
        {
            foreach (IExcelTestTopic topic in TestTopics)
                ExecuteTopic(topic);
        }

        /// <summary>
        /// Execution of all the tests declared on the test topics declared on this class (property 'TestTopics').
        ///<br/>Invoke by double-clicking on the template button 'ExecuteTopic All TestTopics' on the template 'Main' declared on the sheet 'Dashboard Templates'.
        /// </summary>
        public void ExecuteTopic(IExcelTestTopic topic)
        {
            Status = "Executing ...";
            topic.InitTestsStatus();
            ETKExcel.ExcelApplication.PostAsynchronousAction(() =>
                    {
                        try
                        {
                            topic.ExecuteTests();
                        }
                        finally
                        {
                            Status = string.Empty;
                            ExcelInterop.Worksheet dashBoardSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, "Dashboard");
                            if (dashBoardSheet != null)
                                ((ExcelInterop._Worksheet) dashBoardSheet).Activate();
                       }
                    });
        }
        #endregion
    }
}
