namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalMonoHeaderAndFooterTests : ExcelTestTopic
    {
        public BasicVerticalMonoHeaderAndFooterTests(IExcelTestsManager testManager)
               : base(testManager, "Tests on a basic template (without linked templates) with a one line header and a one line footer")
        {
            Tests.Add(new TestCompleteView(this));
            Tests.Add(new TestViewParts(this));
        }

        override protected void RealInit()
        {
            CreateView("VerticalMonoHeaderAndFooter", "BasicTemplates1", "BasicVerticalMonoHeaderAndFooter");

            View.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
