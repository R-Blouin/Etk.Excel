namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestBody : ExcelTest
    {
        public TestBody(IExcelTemplateView view) : base(view, "Check Body")
        { }

        override protected void RealExecute()
        {
            Success = true;
        }
    }
}
