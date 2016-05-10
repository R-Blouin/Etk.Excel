namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalNoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Etk.Excel.BindingTemplates.Views;

    class TestRendering : ExcelTest
    {
        public TestRendering(): base("Check Rendering")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {
            if (view.RenderedArea == null)
                ErrorMessages.Add("Rendered area must not be null");
            else
            {
                if (view.RenderedArea.Width != 4 || view.RenderedArea.Height != 4)
                    ErrorMessages.Add("Rendered area must be 4*4");

                if (view.RenderedRange[1, 1].Value != 1)
                    ErrorMessages.Add("First cell must contains '1'");

                if (view.RenderedRange[4, 4].Value != "Fourth Shop Reception Phone number")
                    ErrorMessages.Add("Last cell must contains 'Fourth Shop Reception Phone number'");
            }
        }
    }
}
