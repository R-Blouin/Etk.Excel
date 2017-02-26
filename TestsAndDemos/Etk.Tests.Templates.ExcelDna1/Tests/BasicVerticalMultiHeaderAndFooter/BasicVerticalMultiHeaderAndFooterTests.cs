namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMultiHeaderAndFooter
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalMultiHeaderAndFooterTests : ExcelTestTopic
    {
        public BasicVerticalMultiHeaderAndFooterTests(IExcelTestsManager testManager)
               : base(testManager, 2, "Tests on a basic template (without linked templates) with a 2 lines header and 2 lines footer", "VerticalMultiHeaderAndFooter")
        {
            Tests.Add(new TestCompleteView(this));
            Tests.Add(new TestViewParts(this));
        }

        override protected void RenderViews()
        {
            CreateViews("BasicTemplates1", "BasicVerticalMultiHeaderAndFooter");

            TopicView.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(TopicView);
        }
    }
}
