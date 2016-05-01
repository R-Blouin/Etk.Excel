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
            if (View.RenderedArea == null)
                ErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (View.RenderedArea.Width != 4 || View.RenderedArea.Height != 4)
                    ErrorMessages.Add("Rendered area must be 4*4");

                if (View.RenderedRange[1, 1].Value != 1)
                    ErrorMessages.Add("First cell must contains '1'");

                if (View.RenderedRange[4, 4].Value != "Fourth Shop Reception Phone number")
                    ErrorMessages.Add("Last cell must contains 'Fourth Shop Reception Phone number'");
            }
        }
    }
}
