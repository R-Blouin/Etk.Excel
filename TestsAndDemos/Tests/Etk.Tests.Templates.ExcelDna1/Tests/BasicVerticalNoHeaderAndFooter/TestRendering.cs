namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Etk.Excel.BindingTemplates.Views;

    class TestRendering : ExcelTest
    {
        public TestRendering(IExcelTemplateView view): base(view, "Check Rendering")
        { }

        override protected void RealExecute()
        {
            Success = true;
            if (View.RenderedArea == null || View.RenderedArea.Width != 4 || View.RenderedArea.Height != 4)
                Success = false;

            if (Success && View.RenderedRange[1, 1].Value != 1)
                Success = false;

            if (Success && View.RenderedRange[4, 4].Value != "Founth Shop Reception Phone number")
                Success = false;
        }
    }
}
