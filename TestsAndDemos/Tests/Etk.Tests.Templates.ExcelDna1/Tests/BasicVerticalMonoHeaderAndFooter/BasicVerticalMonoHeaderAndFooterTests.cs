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
        public BasicVerticalMonoHeaderAndFooterTests() : base("Tests on a basic template (without linked templates) with a one line header and a one line footer")
        {
            Tests.Add(new TestCompleteView());
            Tests.Add(new TestViewParts());
            Tests.Add(new TestDoubleRendering());
        }

        override protected void RealInit()
        {
            CreateView("VerticalMonoHeaderAndFooter", "BasicTemplates1", "BasicVerticalMonoHeaderAndFooter");

            View.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
