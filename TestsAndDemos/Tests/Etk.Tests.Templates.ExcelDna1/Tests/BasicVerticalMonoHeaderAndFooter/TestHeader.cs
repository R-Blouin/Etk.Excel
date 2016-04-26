namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestHeader : ExcelTest
    {
        public TestHeader(IExcelTemplateView view): base(view, "Check Header")
        { }

        override protected void RealExecute()
        {
            Success = true;
        }
    }
}
