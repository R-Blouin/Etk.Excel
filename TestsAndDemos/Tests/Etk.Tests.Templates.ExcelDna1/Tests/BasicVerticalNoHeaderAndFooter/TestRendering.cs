namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter
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
                if (view.RenderedArea.Width != 4 || view.RenderedArea.Height != 4)
                    StepsErrorMessages.Add("Rendered area must be 4*4");

                if (view.RenderedRange[1, 1].Value != 1)
                    StepsErrorMessages.Add("Cell [1, 1] must contains '1'");
                if (view.RenderedRange[1, 2].Value != "First Shop")
                    StepsErrorMessages.Add("Cell [1, 2] must contains 'First Shop'");
                if (view.RenderedRange[1, 4].Value != "First Shop Reception Phone number")
                    StepsErrorMessages.Add("Cell [1, 4] must contains 'First Shop Reception Phone number'");

                if (view.RenderedRange[2, 1].Value != 2)
                    StepsErrorMessages.Add("Cell [2, 1] must contains '2'");
                if (view.RenderedRange[2, 2].Value != "Second Shop")
                    StepsErrorMessages.Add("Cell [2, 2] must contains 'Second Shop'");
                if (view.RenderedRange[2, 4].Value != "Second Shop Reception Phone number")
                    StepsErrorMessages.Add("Cell [2, 4] must contains 'Second Shop Reception Phone number'");

                if (view.RenderedRange[4, 1].Value != 4)
                    StepsErrorMessages.Add("Cell [4, 1] must contains '4'");
                if (view.RenderedRange[4, 2].Value != "Fourth Shop")
                    StepsErrorMessages.Add("Cell [4, 2] must contains 'Fourth Shop'");
                if (view.RenderedRange[4, 4].Value != "Fourth Shop Reception Phone number")
                    StepsErrorMessages.Add("Cell [4, 4] must contains 'Fourth Shop Reception Phone number'");
            }
        }
    }
}
