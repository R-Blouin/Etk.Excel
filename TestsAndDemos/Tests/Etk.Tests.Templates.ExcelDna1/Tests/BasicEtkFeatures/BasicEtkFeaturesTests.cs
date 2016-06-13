namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicEtkFeatures
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Etk.Excel.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;
    using Etk.Excel;
    using Etk.Tests.Data.Shops;
    
    class BasicEtkFeaturesTests : ExcelTestTopic
    {
        public BasicEtkFeaturesTests(IExcelTestsManager testManager)
            : base(testManager, "Test basic ETK features")
        {
            Tests.Add(new TestDoubleRendering(this));
        }

        override protected void RealInit()
        {
            CreateView("BasicEtkFeatures", "BasicTemplates1", "BasicVerticalNoHeaderAndFooter");
            View.SetDataSource(ShopManager.GetShops());
            ETKExcel.TemplateManager.Render(View);
        }
    }
}
