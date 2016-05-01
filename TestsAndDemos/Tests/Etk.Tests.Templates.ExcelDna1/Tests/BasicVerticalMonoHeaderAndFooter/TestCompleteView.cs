namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestCompleteView : ExcelTest
    {
        public TestCompleteView(IExcelTemplateView view) : base(view, "Check complete rendering")
        { }

        override protected void RealExecute()
        {
            if (View.RenderedArea == null)
                ErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (View.RenderedArea == null || View.RenderedArea.Width != 4 || View.RenderedArea.Height != 6)
                    ErrorMessages.Add("Rendered area must be 4*6");

                if (View.RenderedRange[1, 1].Value != "ID")
                    ErrorMessages.Add("First cell must contains 'ID'");

                if (View.RenderedRange[6, 1].Value != "Shops")
                    ErrorMessages.Add("First cell of last row must contains 'Shops'");
            }
        }
    }
}
