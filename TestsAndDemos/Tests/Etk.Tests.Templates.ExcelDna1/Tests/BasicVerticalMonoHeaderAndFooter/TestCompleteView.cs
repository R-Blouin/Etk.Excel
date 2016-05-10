namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestCompleteView : ExcelTest
    {
        public TestCompleteView() : base("Check complete rendering")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {
            if (view.RenderedArea == null)
                ErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (view.RenderedArea == null || view.RenderedArea.Width != 4 || view.RenderedArea.Height != 6)
                    ErrorMessages.Add("Rendered area must be 4*6");

                if (view.RenderedRange[1, 1].Value != "ID")
                    ErrorMessages.Add("First cell must contains 'ID'");

                if (view.RenderedRange[6, 1].Value != "Shops")
                    ErrorMessages.Add("First cell of last row must contains 'Shops'");
            }
        }
    }
}
