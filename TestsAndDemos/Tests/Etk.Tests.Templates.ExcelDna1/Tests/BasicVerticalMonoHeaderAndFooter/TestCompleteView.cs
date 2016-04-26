namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestCompleteView : ExcelTest
    {
        public TestCompleteView(IExcelTemplateView view) : base(view, "Check complete rendering")
        { }

        override protected void RealExecute()
        {
            Success = true;
        }
    }
}
