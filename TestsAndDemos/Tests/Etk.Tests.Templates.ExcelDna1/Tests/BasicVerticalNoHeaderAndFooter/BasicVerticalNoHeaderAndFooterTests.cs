namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Etk.Excel.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;
    using Etk.Excel;
    using Etk.Tests.Data.Shops;
    
    class BasicVerticalNoHeaderAndFooterTests : ExcelTests
    {
        public BasicVerticalNoHeaderAndFooterTests() : base("Render a basic template (without linked templates) without header and footer")
        {}

        override protected void RealInit()
        {
            CreateView("VerticalNoHeaderAndFooter", "BasicTemplates1", "BasicVerticalNoHeaderAndFooter");
            TestsList.Add(new TestRendering(View));
        }

        override protected void RenderViews()
        {
            View.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
