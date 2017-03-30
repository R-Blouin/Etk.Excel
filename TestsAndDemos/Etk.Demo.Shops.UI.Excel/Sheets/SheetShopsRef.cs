using System.Windows;
using Etk.Demo.Shops.UI.Common.ViewModels;
using Etk.Demo.Shops.UI.Excel.Panels;
using Etk.Excel;
using Etk.Excel.BindingTemplates.Views;
using ExcelDna.Integration.CustomUI;

namespace Etk.Demo.Shops.UI.Excel.Sheets
{
    class SheetShopsRef
    {
        private static SheetShopsRef shopsSheet;
        public static SheetShopsRef Instance
        {
            get
            {
                if (shopsSheet == null)
                    shopsSheet = new SheetShopsRef(new ShopsViewModel());
                return shopsSheet;
            }
        }

        private readonly ShopsViewModel viewModel;
        private readonly CustomTaskPane shopsTaskPane;
        private IExcelTemplateView view;

        #region .ctors and factories
        private SheetShopsRef(ShopsViewModel viewModel)
        {
            this.viewModel = viewModel;

            // Create and display a taskpane => to see the interaction between data in Wpf UI and ETK templates
            shopsTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(ShopsPanel), "Shops Panel");
            shopsTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;

            (shopsTaskPane.ContentControl as ShopsPanel).SetViewModel(viewModel);
            shopsTaskPane.Visible = true;
        }
        #endregion

        #region 
        /// <summary> Create and render view</summary>
        public void RenderViews()
        {
            if (view != null)
            {
                viewModel.PropertyChanged -= OnViewPropertiesChanged;
                ETKExcel.TemplateManager.RemoveView(view);
            }

            view = ETKExcel.TemplateManager.AddView("TemplatesShops_Ref", "Main", "Shops", "B2");
            // Inject the data source
            view.SetDataSource(viewModel.ShopsToDisplay);
            // RenderView the sheet
            view.Render();

            view.ViewSheetIsActivated += notUsedParameter => shopsTaskPane.Visible = true;
            view.ViewSheetIsDeactivated += notUsedParameter => shopsTaskPane.Visible = false;

            viewModel.PropertyChanged += OnViewPropertiesChanged;
        }

        private void OnViewPropertiesChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ShopsToDisplay")
                RenderViews();
        }
        #endregion

        public static void DisplayName(CustomerViewModel customer)
        {
            MessageBox.Show($"{customer.Customer.Forename} {customer.Customer.Surname}");
        }
    }
}
