﻿using Etk.BindingTemplates.Views;
using Etk.Excel;
using Etk.Excel.Application;
using Etk.Excel.MvvmBase;
using Etk.Tests.Templates.ExcelDna1.Tests;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Etk.Tests.Templates.ExcelDna1
{
    interface IExcelTestsManager
    {
        void ExecuteTopics(IEnumerable<IExcelTestTopic> topicq);
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

        public void ExecuteTopics(IEnumerable<IExcelTestTopic> topics)
        {
            Status = "Executing ...";
            Action action = new Action(() =>
            {
                using (FreezeExcel freeExcel_ = new FreezeExcel())
                {
                    foreach (IExcelTestTopic topic in topics)
                        topic.ExecuteTests();
                }
            });
            ETKExcel.ExcelApplication.PostAsynchronousActions(new[] { action }, () => Status = string.Empty);
        }

        #endregion
    }
}
