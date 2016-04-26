namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Tests.Templates.ExcelDna1.Extensions;
    using Microsoft.Office.Interop.Excel;

    abstract class ExcelTests : IExcelTests
    {
        #region properties and attributes
        private Worksheet templatesSheet = null;
        private Worksheet viewSheet = null;

        public IExcelTemplateView View
        { get; private set; }

        protected List<IExcelTest> TestsList
        { get; private set; }

        public string Description
        { get; private set; }

        public IEnumerable<IExcelTest> Tests
        { get { return TestsList; } }

        public bool InitSuccessful
        { get; private set; }

        public string Exception
        { get; private set; }
        #endregion

        #region .ctors
        protected ExcelTests(string description)
        {
            Description = description;
            TestsList = new List<IExcelTest>();
        }
        #endregion

        #region pubic methods
        public void Init()
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
        }

        public void Execute()
        {
            if (! InitSuccessful)
                return;

            try
            {
                try
                {
                    RenderViews();
                }
                catch (Exception ex)
                {
                    throw new Exception("Render views failed", ex);
                }
                TestsList.ForEach(t => t.Execute());
            }
            catch (Exception ex)
            {
                Exception = ex.ToString("Execution failed");
            }
        }
        #endregion

        #region protected methods
        abstract protected void RenderViews();
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
        #endregion
    }
}
