namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Extensions;
    using Microsoft.Office.Interop.Excel;
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    abstract class ExcelTestTopic : ViewModelBase, IExcelTestTopic
    {
        #region properties and attributes
        private IExcelTestsManager testManager;
        private Worksheet templatesSheet = null;
        private Worksheet viewSheet = null;
        private bool initDone;

        public IExcelTemplateView View
        { get; private set; }

        public string Description
        { get; private set; }

        public List<IExcelTest> Tests
        { get; private set; } 

        private bool initSuccessful;
        public bool InitSuccessful
        {
            get { return initSuccessful; }
            private set
            {
                initSuccessful = value;
                OnPropertyChanged("InitSuccessful");
            }
        }

        private string exception;
        public string Exception
        {
            get { return exception; }
            private set
            {
                exception = value;
                OnPropertyChanged("Exception");
            }
        }
        #endregion

        #region .ctors
        protected ExcelTestTopic(IExcelTestsManager testManager, string description)
        {
            this.testManager = testManager;
            Description = description;
            Tests = new List<IExcelTest>();
        }
        #endregion

        #region pubic methods
        public void InitTestsStatus()
        {
            Tests.ForEach(t => t.InitTestStatus());
        }

        /// <summary>
        /// ExecuteTopic all th tests declared on this topic (property 'Tests').
        /// Invoke by double-clicking on the template button 'ExecuteTopic' on the template 'TestTopics' declared on the sheet 'Dashboard Templates'
        /// </summary>
        public void Execute()
        {
            testManager.ExecuteTopic(this);
        }

        public void ExecuteTests()
        {
            if (!initDone)
                Init();

            if (!InitSuccessful)
                return;

            try
            {
                Tests.ForEach(t => t.Execute(View));
            }
            catch (Exception ex)
            {
                Exception = ex.ToString("Execution failed");
            }
        }

        public int GetNumberOfTests()
        {
            return Tests.Count;
        }
        #endregion

        #region protected methods
        abstract protected void RealInit();

        protected void CreateView(string destinationSheetName, string templateSheetName, string templateName)
        {
            if (View != null)
                ETKExcel.TemplateManager.RemoveView(View);

            templatesSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, templateSheetName);
            viewSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, destinationSheetName);
            if(viewSheet == null)
            {
                Workbook workbook = templatesSheet.Parent;
                Sheets sheets = workbook.Sheets;
                Worksheet lastSheets = workbook.Sheets[sheets.Count];

                viewSheet = workbook.Worksheets.Add(Type.Missing, lastSheets); 
                viewSheet.Name = destinationSheetName;

                Marshal.ReleaseComObject(lastSheets);
                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(workbook);
            }
            else
            {
                Range usedRange = viewSheet.UsedRange;
                if(usedRange != null)
                    usedRange .Clear();
                usedRange  = null;
            }

            Range firstRange = viewSheet.Range["B2"];
            View = ETKExcel.TemplateManager.AddView(templatesSheet, templateName, viewSheet, firstRange);
            firstRange = null;
        }
        #endregion

        #region private methods
        private void Init()
        {
            try
            {
                RealInit();
                InitSuccessful = true;
            }
            catch (Exception ex)
            {
                InitSuccessful = false;
                Exception = ex.ToString("Initialization failed");
            }
            finally
            {
                initDone = true;
            }
        }
        #endregion
    }
}
