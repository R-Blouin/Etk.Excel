namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;
    
    class BasicVerticalNoHeaderAndFooterTests : ExcelTestTopic
    {
        public BasicVerticalNoHeaderAndFooterTests(IExcelTestsManager testManager) 
               : base(testManager, 0, "Tests on a basic template without linked templates and without header or footer", "VerticalNoHeaderAndFooter")
        {
            Tests.Add(new TestRendering(this));
        }

        override protected void RenderViews()
        {
            CreateViews("BasicTemplates1", "BasicVerticalNoHeaderAndFooter");
            TopicView.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(TopicView);
        }
    }
}
