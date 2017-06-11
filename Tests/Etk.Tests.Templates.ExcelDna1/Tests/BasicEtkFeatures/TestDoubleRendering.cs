namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicEtkFeatures
{
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;

    class TestDoubleRendering : ExcelTest
    {
        public TestDoubleRendering(IExcelTestTopic parent): base(parent, "Try to render after a first rendering with no clear or SetDataSource between")
        {}

        override protected void RealExecute(IExcelTemplateView view)
        {
            ETKExcel.TemplateManager.Render(view);
        }
    }
}
