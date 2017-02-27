using Etk.Demo.Shops.UI.Common.Controls.ViewModels;
using Etk.Demo.Shops.UI.Excel.Panels;
using Etk.Excel;
using Etk.Excel.BindingTemplates.Views;
using ExcelDna.Integration.CustomUI;

namespace Etk.Demo.Shops.UI.Excel.Sheets
{
    class SheetShops
    {
        private IExcelTemplateView view;
        private ShopsViewModel viewModel;
        private CustomTaskPane shopsTaskPane;


        #region .ctors and factories

        private SheetShops(ShopsViewModel viewModel)
        {
            this.viewModel = viewModel;
        }

        public static void CreateAndActivateDashBoard()
        {
            SheetShops shopsSheet = new SheetShops(new ShopsViewModel());
            shopsSheet.CreateShopsPanel();
            shopsSheet.CreateAndRender();
        }
        #endregion

        #region 
        private void CreateShopsPanel()
        {
            // Create and display a taskpane => to see the interaction between data in Wpf UI and ETK templates
            shopsTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(ShopsPanel), "Shops Panel");
            shopsTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;

            (shopsTaskPane.ContentControl as ShopsPanel).SetViewModel(viewModel);
            shopsTaskPane.Visible = true;
        }

        /// <summary> Create and render the dashboard view</summary>
        private void CreateAndRender()
        {
            view = ETKExcel.TemplateManager.AddView("TemplatesShops", "Main", "Shops", "B2");
            // Inject the data source
            view.SetDataSource(viewModel.ShopsToDisplay);
            // RenderView the sheet
            ETKExcel.TemplateManager.Render(view);

            view.ViewSheetIsActivated += (notUsedParameter) =>
                                {
                                    //using (FreezeExcel freeExcel = new FreezeExcel())
                                    {
                                        shopsTaskPane.Visible = true;
                                    }
                                };
            view.ViewSheetIsDeactivated += (notUsedParameter) =>
                                {
                                    //using (FreezeExcel freeExcel = new FreezeExcel())
                                    {
                                        shopsTaskPane.Visible = false;
                                    }
                                };
        }
        #endregion
    }
}
