namespace Etk.Tests.Templates.ExcelDna1
{
    using Etk.BindingTemplates.Views;
    using Etk.Excel;
    using Excel.Application;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Tests;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    interface IExcelTestsManager
    {
        void ExecuteTopic(IExcelTestTopic topic);
    }

    class ExcelTestsManager : ViewModelBase, IExcelTestsManager 
    {
        #region attributes and properties 
        public IEnumerable<IExcelTestTopic> TestTopics { get; private set; }

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
            IEnumerable<Type> types = Assembly.GetExecutingAssembly().GetTypes().Where(t => t.IsSubclassOf(typeof(ExcelTestTopic)));
            TestTopics = types.Select(t => Activator.CreateInstance(t, new[] { this }) as IExcelTestTopic)
                              .OrderBy(t => t.Id)
                              .ThenBy(t => t.Description)               
                              .ToArray();

            using (FreezeExcel freeExcel = new FreezeExcel())
            {
                foreach (IExcelTestTopic topic in TestTopics)
                    topic.Init();
           }
        }
        #endregion

        #region public methods
        public static void SetSearchValue(ITemplateView concernedView, IExcelTestTopic topic)
        {
            concernedView.SearchValue = concernedView.SearchValue == topic.Description ? null : topic.Description;
            concernedView.ExecuteSearch();
        }

        /// <summary>
        /// Execution of all the tests of the test topics.
        ///<br/>Invoke by double-clicking on the template button 'ExecuteTopic All TestTopics' on the template 'Main' declared on the sheet 'Dashboard Templates'.
        /// </summary>
        public void Execute()
        {
            ExecuteTopics(TestTopics);
        }

        /// <summary>
        /// Execution of all the tests of the current test topic declared on this class (property 'TestTopics').
        ///<br/>Invoke by double-clicking on the template button 'ExecuteTopic All TestTopics' on the template 'Main' declared on the sheet 'Dashboard Templates'.
        /// </summary>
        public void ExecuteTopic(IExcelTestTopic topic)
        {
               ExecuteTopics(new[] { topic });
        }
        #endregion

        #region private methods
        private void ExecuteTopics(IEnumerable<IExcelTestTopic> topics)
        {
            Status = "Executing ...";
            Action action = new Action(() => 
                            {
                                using (FreezeExcel freeExcel_ = new FreezeExcel())
                                {
                                    foreach(IExcelTestTopic topic in topics)
                                    {
                                        topic.InitTestsStatus();
                                        topic.ExecuteTests();
                                    }
                                }
                            });
            ETKExcel.ExcelApplication.PostAsynchronousActions(new[] { action }, () => Status = string.Empty);
        }
        #endregion
    }
}
