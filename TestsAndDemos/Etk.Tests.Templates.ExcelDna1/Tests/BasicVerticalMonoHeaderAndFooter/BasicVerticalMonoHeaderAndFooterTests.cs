namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalMonoHeaderAndFooterTests : ExcelTestTopic
    {
        public BasicVerticalMonoHeaderAndFooterTests(IExcelTestsManager testManager)
               : base(testManager, 1, "Tests on a basic template (without linked templates) with a one line header and a one line footer", "VerticalMonoHeaderAndFooter")
        {
            Tests.Add(new TestCompleteView(this));
            Tests.Add(new TestViewParts(this));
        }

        override protected void RenderViews()
        {
            CreateViews("BasicTemplates1", "BasicVerticalMonoHeaderAndFooter");

            TopicView.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(TopicView);
        }
    }
}
