namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicVerticalMonoHeaderAndFooter
{
    using System;
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;

    class TestViewParts : ExcelTest
    {
        public TestViewParts() : base("Check the rendering of the template parts")
        { }

        override protected void RealExecute(IExcelTemplateView view)
        {

            ExcelTemplateView excelView = view as ExcelTemplateView;

            if (excelView.RenderedArea == null || excelView.Renderer == null)
            {
                ErrorMessages.Add("Rendered area must not be null");
                return;
            }

            // Header
            if (excelView.Renderer.HeaderPartRenderer == null || excelView.Renderer.HeaderPartRenderer.RenderedArea == null)
                ErrorMessages.Add("Header rendered area must not be null");
            else
            {
                if(excelView.Renderer.HeaderPartRenderer.Width != 4 || excelView.Renderer.HeaderPartRenderer.Height != 1)
                    ErrorMessages.Add("Header rendered area must 1*1");
                if (excelView.Renderer.HeaderPartRenderer.RenderedRange[1, 1].Value != "ID")
                    ErrorMessages.Add("First cell must contains 'ID'");
            }

            // Body
            if (excelView.Renderer.BodyPartRenderer == null || excelView.Renderer.BodyPartRenderer.RenderedArea == null)
                ErrorMessages.Add("Body rendered area must not be null");
            else
            {
                if (excelView.Renderer.BodyPartRenderer.Width != 4 || excelView.Renderer.BodyPartRenderer.Height != 4)
                    ErrorMessages.Add("Body Rendered area must be 4*4");
                if (excelView.Renderer.BodyPartRenderer.RenderedRange[1, 1].Value != 1)
                    ErrorMessages.Add("First cell must contains '1'");
                if (excelView.Renderer.BodyPartRenderer.RenderedRange[4, 4].Value != "Fourth Shop Reception Phone number")
                    ErrorMessages.Add("Last cell must contains 'Fourth Shop Reception Phone number'");
            }

            // Footer
            if (excelView.Renderer.FooterPartRenderer == null || excelView.Renderer.FooterPartRenderer.RenderedArea == null)
                ErrorMessages.Add("Footer rendered area must not be null");
            else
            {
                if (excelView.Renderer.FooterPartRenderer.Width != 4 || excelView.Renderer.HeaderPartRenderer.Height != 1)
                    ErrorMessages.Add("Footer rendered area must 4*1");
                if (excelView.Renderer.FooterPartRenderer.RenderedRange[1, 1].Value != "Shops")
                    ErrorMessages.Add("First cell of last row must contains 'Shops'");
            }
        }
    }
}
