namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalWithNothingElseThanALinkedTemplate
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalWithNothingElseThanALinkedTemplateTests : ExcelTestTopic
    {
        public BasicVerticalWithNothingElseThanALinkedTemplateTests(IExcelTestsManager testManager) : base(testManager, "Tests on a basic template with nothing else than one linked template")
        {
            Tests.Add(new TestRendering(this));
        }

        override protected void RealInit()
        {
            CreateView("VerticalWithOnlyOneLink", "BasicTemplates1", "BasicVerticalWithNothingElseThanALinkedTemplate");
            
            View.SetDataSource(ProductsManager.Instance);
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
