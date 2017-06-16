namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicEtkFeatures
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;
    
    class BasicEtkFeaturesTests : ExcelTestTopic
    {
        public BasicEtkFeaturesTests(IExcelTestsManager testManager)
               : base(testManager, 100, "Tests basic ETK features", "BasicEtkFeatures")
        {
            Tests.Add(new TestDoubleRendering(this));
        }

        override protected void RenderViews()
        {
            CreateViews("BasicTemplates1", "BasicVerticalNoHeaderAndFooter");
            TopicView.SetDataSource(ShopManager.Shops);
            ETKExcel.TemplateManager.Render(TopicView);
        }
    }
}
