//using System.Windows;
//using Etk.Demo.Shops.UI.Common.ViewModels;
//using Etk.Demo.Shops.UI.Excel.Panels;
//using Etk.Excel;
//using Etk.Excel.BindingTemplates.Views;
//using ExcelDna.Integration.CustomUI;

//namespace Etk.Demo.Shops.UI.Excel.Sheets
//{
//    class SheetShops
//    {
//        private readonly ShopsViewModel viewModel;
//        private IExcelTemplateView view;

//        private static SheetShops shopsSheet;
//        public static SheetShops Instance
//        {
//            get
//            {
//                if (shopsSheet == null)
//                    shopsSheet = new SheetShops(new ShopsViewModel());
//                return shopsSheet;
//            }
//        }

//        #region .ctors and factories
//        private SheetShops(ShopsViewModel viewModel)
//        {
//            this.viewModel = viewModel;
//        }
//        #endregion

//        /// <summary> Create and render view</summary>
//        public void RenderViews()
//        {
//            if (view != null)
//            {
//                viewModel.PropertyChanged -= OnViewPropertiesChanged;
//                ETKExcel.TemplateManager.RemoveView(view);
//            }

//            view = ETKExcel.TemplateManager.AddView("TemplatesShops", "Main", "Shops", "B2");
//            // Inject the data source
//            view.SetDataSource(viewModel.ShopsToDisplay);
//            // RenderView the sheet
//            view.Render();

//            viewModel.PropertyChanged += OnViewPropertiesChanged;
//        }

//        private void OnViewPropertiesChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
//        {
//            if (e.PropertyName == "ShopsToDisplay")
//                RenderViews();
//        }

//        public static void DisplayName(CustomerViewModel customer)
//        {
//            MessageBox.Show($"{customer.Customer.Forename} {customer.Customer.Surname}");
//        }
//    }
//}
