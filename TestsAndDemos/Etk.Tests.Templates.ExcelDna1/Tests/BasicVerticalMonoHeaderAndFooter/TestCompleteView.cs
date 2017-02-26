namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestCompleteView : ExcelTest
    {
        public TestCompleteView(IExcelTestTopic parent): base(parent, "Control the size of the rendered area and the value of some cells")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {
            if (view.RenderedArea == null)
                StepsErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (view.RenderedArea == null || view.RenderedArea.Width != 4 || view.RenderedArea.Height != 6)
                    StepsErrorMessages.Add("Rendered area must be 4*6");

                if (view.RenderedRange[1, 1].Value != "ID")
                    StepsErrorMessages.Add("First cell must contains 'ID'");

                if (view.RenderedRange[1, 4].Value != "Reception Phone Number")
                    StepsErrorMessages.Add("Cells [1, 3] must contains 'Reception Phone Number'");

                if (view.RenderedRange[2, 1].Value != 1)
                    StepsErrorMessages.Add("Cells [2, 1] must contains '1'");

                if (view.RenderedRange[2, 4].Value != "First Shop Reception Phone number")
                    StepsErrorMessages.Add("Cells [2, 3] must contains 'First Shop Reception Phone number'");

                if (view.RenderedRange[4, 2].Value != "Third Shop")
                    StepsErrorMessages.Add("Cells [4, 2] must contains 'Third Shop'");

                if (view.RenderedRange[4, 4].Value != "Third Shop Reception Phone number")
                    StepsErrorMessages.Add("Cells [2, 3] must contains 'Third Shop Reception Phone number'");

                if (view.RenderedRange[6, 1].Value != "Shops")
                    StepsErrorMessages.Add("First cell of last row must contains 'Shops'");
            }
        }
    }
}
