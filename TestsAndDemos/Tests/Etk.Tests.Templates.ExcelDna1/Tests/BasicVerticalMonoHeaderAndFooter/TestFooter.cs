namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestFooter : ExcelTest
    {
        public TestFooter(IExcelTemplateView view) : base(view, "Check Footer")
        { }

        override protected void RealExecute()
        {
            Success = true;
        }
    }
}
