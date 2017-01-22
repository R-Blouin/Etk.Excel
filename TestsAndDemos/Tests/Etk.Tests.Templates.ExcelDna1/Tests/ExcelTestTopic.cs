namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Extensions;
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    abstract class ExcelTestTopic : ViewModelBase, IExcelTestTopic
    {
        #region properties and attributes
        private IExcelTestsManager testManager;
        private ExcelInterop.Worksheet templatesSheet = null;
        private ExcelInterop.Worksheet viewsOwnerSheet = null;
        private bool renderDone;

        public IExcelTemplateView TopicView { get; private set; }
        public IExcelTemplateView GoBackView { get; private set; }

        public int Id { get; private set; }

        public string Description { get; private set; }

        public List<IExcelTest> Tests { get; private set; }

        public string DestinationSheetName { get; private set; }

        private bool renderSuccessful;
        public bool RenderSuccessful
        {
            get { return renderSuccessful; }
            private set
            {
                renderSuccessful = value;
                OnPropertyChanged("RenderSuccessful");
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
        protected ExcelTestTopic(IExcelTestsManager testManager, int id, string description, string destinationSheetName)
        {
            this.testManager = testManager;
            Id = id;
            Description = description;
            DestinationSheetName = destinationSheetName;
            Tests = new List<IExcelTest>();
        }
        #endregion

        #region pubic methods
        public void Init()
        {
            try
            {
                viewsOwnerSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, DestinationSheetName);

                // Create the destination sheet
                ExcelInterop.Workbook workbook = ETKExcel.ExcelApplication.Application.ActiveWorkbook;
                ExcelInterop.Sheets sheets = workbook.Sheets;
                ExcelInterop.Worksheet lastSheet = workbook.Sheets[sheets.Count];
                ExcelInterop.Worksheet firstSheet = workbook.Sheets[1];

                viewsOwnerSheet = workbook.Worksheets.Add(Type.Missing, lastSheet);
                viewsOwnerSheet.Name = DestinationSheetName;
                viewsOwnerSheet.Visible = ExcelInterop.XlSheetVisibility.xlSheetHidden;

                firstSheet.Activate();

                Marshal.ReleaseComObject(firstSheet);
                Marshal.ReleaseComObject(lastSheet);
                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(workbook);
                // End create the destination sheet

                // Create the 'GoBackToDashboard' view
                GoBackView = ETKExcel.TemplateManager.AddView("Dashboard Templates", "GoBackToDashboard", DestinationSheetName, "A1");
                GoBackView.SetDataSource(new GoBackToDashBoardManager());
                GoBackView.Render();
                // End create the 'GoBackToDashboard' view
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Init Topics failed:{0}", ex.Message), ex);
            }
        }

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
            if (!renderDone)
                Render();

            if (!RenderSuccessful)
                return;

            try
            {
                Tests.ForEach(t => t.Execute(TopicView));
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
        abstract protected void RenderViews();

        protected void CreateViews(string templateSheetName, string templateName)
        {
            CreateTopicView(templateSheetName, templateName);
        }
        #endregion

        #region private methods
        private void Render()
        {
            try
            {
                RenderViews();
                RenderSuccessful = true;
            }
            catch (Exception ex)
            {
                RenderSuccessful = false;
                Exception = ex.ToString("Render failed");
            }
            finally
            {
                renderDone = true;
            }
        }
 
        private void CreateTopicView(string templateSheetName, string templateName)
        {
            try
            {
                if (TopicView == null)
                {
                    templatesSheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(ETKExcel.ExcelApplication.Application.ActiveWorkbook, templateSheetName);
                    viewsOwnerSheet.Visible = ExcelInterop.XlSheetVisibility.xlSheetVisible;
                }
                else
                    ETKExcel.TemplateManager.RemoveView(TopicView);

                ExcelInterop.Range firstRange = viewsOwnerSheet.Range["B3"];
                TopicView = ETKExcel.TemplateManager.AddView(templatesSheet, templateName, viewsOwnerSheet, firstRange);
                firstRange = null;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Cannot create topic view:{0}", ex.Message), ex);
            }
        }
        #endregion
    }
}
