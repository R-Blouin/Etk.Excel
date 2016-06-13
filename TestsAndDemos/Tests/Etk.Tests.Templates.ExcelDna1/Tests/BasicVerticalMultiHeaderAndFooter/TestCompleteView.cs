namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMultiHeaderAndFooter
{
    using Etk.Excel.BindingTemplates.Views;

    class TestCompleteView : ExcelTest
    {
        public TestCompleteView(IExcelTestTopic parent) : base(parent, "Control the size of the rendered area and the value of some cells")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {
            if (view.RenderedArea == null)
                StepsErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (view.RenderedArea == null || view.RenderedArea.Width != 4 || view.RenderedArea.Height != 8)
                    StepsErrorMessages.Add("Rendered area must be 4*8");

                if (view.RenderedRange[1, 1].Value != "Shops")
                    StepsErrorMessages.Add("Cell[1,1] must contains 'Shops'");

                if (view.RenderedRange[2, 1].Value != "ID")
                    StepsErrorMessages.Add("Cell[2,1] must contains 'ID'");
                if (view.RenderedRange[2, 2].Value != "Name")
                    StepsErrorMessages.Add("Cell[2,2] must contains 'Address'");
                if (view.RenderedRange[2, 3].Value != "Address")
                    StepsErrorMessages.Add("Cell[2,3] must contains 'ID'");
                if (view.RenderedRange[2, 4].Value != "Reception Phone Number")
                    StepsErrorMessages.Add("Cell[2,4] must contains 'Reception Phone Number'");

                if (view.RenderedRange[4, 1].Value != 2)
                    StepsErrorMessages.Add("Cell[4,1] must contains '2'");
                if (view.RenderedRange[4, 2].Value != "Second Shop")
                    StepsErrorMessages.Add("Cell[4,2] must contains 'Second Shop'");
                if (view.RenderedRange[4, 3].Value != "2 Shops Road ShopCity")
                    StepsErrorMessages.Add("Cell[4,3] must contains '2 Shops Road ShopCity'");
                if (view.RenderedRange[4, 4].Value != "Second Shop Reception Phone number")
                    StepsErrorMessages.Add("Cell[4,4] must contains 'Second Shop Reception Phone number'");

                if (view.RenderedRange[7, 1].Value != "Footer")
                    StepsErrorMessages.Add("Cell[7,4] must contains 'Shops'");
                if (view.RenderedRange[8, 1].Value != "Shops")
                    StepsErrorMessages.Add("Cell[8,4] must contains 'Shops'");
            }
        }
    }
}
