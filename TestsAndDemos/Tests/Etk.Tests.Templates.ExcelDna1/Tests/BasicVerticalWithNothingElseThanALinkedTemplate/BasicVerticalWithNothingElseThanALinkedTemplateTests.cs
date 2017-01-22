namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalWithNothingElseThanALinkedTemplate
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalWithNothingElseThanALinkedTemplateTests : ExcelTestTopic
    {
        public BasicVerticalWithNothingElseThanALinkedTemplateTests(IExcelTestsManager testManager) 
               : base(testManager, 3,"Tests on a basic template with nothing else than one linked template", "VerticalWithOnlyOneLink")
        {
            Tests.Add(new TestRendering(this));
        }

        override protected void RenderViews()
        {
            CreateViews("BasicTemplates1", "BasicVerticalWithNothingElseThanALinkedTemplate");
            
            TopicView.SetDataSource(ProductsManager.Instance);
            ETKExcel.TemplateManager.Render(TopicView);
        }
    }
}
