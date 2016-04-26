namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Etk.Excel.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;
    using Etk.Excel;
    using Etk.Tests.Data.Shops;

    class BasicVerticalMonoHeaderAndFooterTests : ExcelTests
    {
        public BasicVerticalMonoHeaderAndFooterTests() : base("Render a basic template (without linked templates) with a one line header and a one line footer")
        {}

        override protected void RealInit()
        {
            CreateView("VerticalMonoHeaderAndFooter", "BasicTemplates1", "BasicVerticalMonoHeaderAndFooter");
            TestsList.Add(new TestHeader(View));
            TestsList.Add(new TestBody(View));
            TestsList.Add(new TestFooter(View));
            TestsList.Add(new TestCompleteView(View));
        }

        override protected void RenderViews()
        {
            View.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
