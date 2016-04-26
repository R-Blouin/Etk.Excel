namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Etk.Excel;
    using Etk.Excel.BindingTemplates.Views;
    using Excel = Microsoft.Office.Interop.Excel;
    using Etk.Tests.Data.Shops;

    class BasicViewsTestsManager
    {
        private Excel.Workbook currentWorkbook;

        IExcelTemplateView basicVerticalNoHeaderAndFooter;
        IExcelTemplateView basicVerticalMonoHeaderAndFooter;
        IExcelTemplateView basicVerticalMultiHeaderAndFooter;

        IExcelTemplateView basicHorizontalNoHeaderAndFooter;
        IExcelTemplateView basicHorizontalMonoHeaderAndFooter;
        IExcelTemplateView basicHorizontalMultiHeaderAndFooter;

        IExcelTemplateView basicVerticalWithLinkedTemplates;
        IExcelTemplateView basicHorizontalWithLinkedTemplates;

        public BasicViewsTestsManager(Excel.Workbook currentWorkbook)
        {
            this.currentWorkbook = currentWorkbook;
        }

        public void Execute()
        {
            Excel.Worksheet basicTemplates1Sheet = null;
            Excel.Worksheet basicViewsSheet1 = null;

            try
            {
                basicTemplates1Sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(currentWorkbook, "BasicTemplates1");
                basicViewsSheet1 = ETKExcel.ExcelApplication.GetWorkSheetFromName(currentWorkbook, "BasicViews1");

                IEnumerable<IExcelTemplateView> views = ETKExcel.TemplateManager.GetSheetViews(basicViewsSheet1);
                ETKExcel.TemplateManager.ClearViews(views);
                Excel.Range firstRange = basicViewsSheet1.Range["B2"];

                // Render vertical views
                ////////////////////////
                basicVerticalNoHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicVerticalNoHeaderAndFooter", basicViewsSheet1, firstRange);
                basicVerticalNoHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicVerticalNoHeaderAndFooter);

                firstRange = basicViewsSheet1.Cells[basicVerticalNoHeaderAndFooter.RenderedArea.YPos + basicVerticalNoHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicVerticalMonoHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicVerticalMonoHeaderAndFooter", basicViewsSheet1, firstRange);
                basicVerticalMonoHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicVerticalMonoHeaderAndFooter);

                firstRange = basicViewsSheet1.Cells[basicVerticalMonoHeaderAndFooter.RenderedArea.YPos + basicVerticalMonoHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicVerticalMultiHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicVerticalMultiHeaderAndFooter", basicViewsSheet1, firstRange);
                basicVerticalMultiHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicVerticalMultiHeaderAndFooter);

                //// Render horizontal views
                ////////////////////////////
                firstRange = basicViewsSheet1.Cells[basicVerticalMultiHeaderAndFooter.RenderedArea.YPos + basicVerticalMultiHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicHorizontalNoHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicHorizontalNoHeaderAndFooter", basicViewsSheet1, firstRange);
                basicHorizontalNoHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicHorizontalNoHeaderAndFooter);

                firstRange = basicViewsSheet1.Cells[basicHorizontalNoHeaderAndFooter.RenderedArea.YPos + basicHorizontalNoHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicHorizontalMonoHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicHorizontalMonoHeaderAndFooter", basicViewsSheet1, firstRange);
                basicHorizontalMonoHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicHorizontalMonoHeaderAndFooter);

                firstRange = basicViewsSheet1.Cells[basicHorizontalMonoHeaderAndFooter.RenderedArea.YPos + basicHorizontalMonoHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicHorizontalMultiHeaderAndFooter = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicHorizontalMultiHeaderAndFooter", basicViewsSheet1, firstRange);
                basicHorizontalMultiHeaderAndFooter.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicHorizontalMultiHeaderAndFooter);

                // Render views with linked templates
                /////////////////////////////////////
                firstRange = basicViewsSheet1.Cells[basicHorizontalMultiHeaderAndFooter.RenderedArea.YPos + basicHorizontalMultiHeaderAndFooter.RenderedArea.Height + 2, 2];
                basicVerticalWithLinkedTemplates = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicVerticalTemplateWithLinkedTemplates", basicViewsSheet1, firstRange);
                basicVerticalWithLinkedTemplates.SetDataSource(ShopManager.GetShops());
                ETKExcel.TemplateManager.Render(basicVerticalWithLinkedTemplates);

                //firstRange = basicViewsSheet1.Cells[basicVerticalWithLinkedTemplates.RenderedArea.YPos + basicVerticalWithLinkedTemplates.RenderedArea.Height + 2, 2];
                //basicHorizontalWithLinkedTemplates = ETKExcel.TemplateManager.AddView(basicTemplates1Sheet, "BasicHorizontalTemplateWithLinkedTemplates", basicViewsSheet1, firstRange);
                //basicHorizontalWithLinkedTemplates.SetDataSource(ShopManager.GetShops());
                //ETKExcel.TemplateManager.Render(basicHorizontalWithLinkedTemplates);
            }
            finally
            {
                basicViewsSheet1.Columns.AutoFit();

                if (basicTemplates1Sheet != null)
                   Marshal.ReleaseComObject(basicTemplates1Sheet);
                if (basicViewsSheet1 != null)
                    Marshal.ReleaseComObject(basicViewsSheet1);
            }
        }
    }
}
