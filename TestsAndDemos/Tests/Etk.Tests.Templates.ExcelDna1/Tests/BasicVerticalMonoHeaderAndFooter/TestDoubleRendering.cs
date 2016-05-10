namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Excel;

    class TestDoubleRendering : ExcelTest
    {
        public TestDoubleRendering(): base("Check to render after a first rendering with no clear or SetDataSource between")
        {}

        override protected void RealExecute(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.Render(view);
        }
    }
}
