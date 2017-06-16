namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalWithNothingElseThanALinkedTemplate
{
    using Etk.Excel.BindingTemplates.Views;

    class TestRendering : ExcelTest
    {
        public TestRendering(IExcelTestTopic parent) : base(parent, "Control the size of the rendered area and the value of some cells")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {
            if (view.RenderedArea == null)
                StepsErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (view.RenderedArea.Width != 3 || view.RenderedArea.Height != 10)
                    StepsErrorMessages.Add("Rendered area must be 3*10");

                if (view.RenderedRange[1, 1].Value != 1)
                    StepsErrorMessages.Add("Cell [1, 1] must contains '1'");
                if (view.RenderedRange[1, 2].Value != "Product 1")
                    StepsErrorMessages.Add("Cell [1, 2] must contains 'Product 1'");
                if (view.RenderedRange[1, 3].Value != 11.0)
                    StepsErrorMessages.Add("Cell [1, 3] must contains '11.00'");

                if (view.RenderedRange[5, 1].Value != 5)
                    StepsErrorMessages.Add("Cell [5, 1] must contains '5'");
                if (view.RenderedRange[5, 2].Value != "Product 5")
                    StepsErrorMessages.Add("Cell [5, 2] must contains 'Product 5'");
                if (view.RenderedRange[5, 3].Value != 55.0)
                    StepsErrorMessages.Add("Cell [5, 3] must contains '55.00'");

                if (view.RenderedRange[10, 1].Value != 10)
                    StepsErrorMessages.Add("Cell [10, 1] must contains '10'");
                if (view.RenderedRange[10, 2].Value != "Product 10")
                    StepsErrorMessages.Add("Cell [10, 2] must contains 'Product 10'");
                if (view.RenderedRange[10, 3].Value != 100.0)
                    StepsErrorMessages.Add("Cell [10, 3] must contains '100.00'");
            }
        }
    }
}
