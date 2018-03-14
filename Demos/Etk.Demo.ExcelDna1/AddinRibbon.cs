using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Etk.Excel;
using Etk.Excel.BindingTemplates.Views;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Etk.Demos.Data.Shops;
using Etk.Demos.Data.Shares;

namespace Etk.Demo.ExcelDna1
{
    /// <summary>Manage the application custom ribbons</summary>
    [ComVisible(true)]
    public class AddinRibbon : ExcelRibbon
    {
        private IExcelTemplateView mainCustomersView;
        private IExcelTemplateView mainSharesView;

        /// <summary>Excel Dna method used to build the ribbon</summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            //ShowCustomers();
            //ShowShares();
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Demo.ExcelDna1.Resources.AddinRibbon.xml")))
            {
                string ribbonXml = textReader.ReadToEnd();
                return ribbonXml;
            }
        }

        #region event Handlers
        public void OnRenderAll(IRibbonControl control)
        {
            Worksheet activeSheet = ETKExcel.ExcelApplication.GetActiveSheet();
            if (activeSheet != null)
            {
                OnClearViews(null);
                switch (activeSheet.Name)
                {
                    case "Customers":
                        ShowCustomers();
                    break;
                    case "Shares":
                        ShowShares();
                    break;
                }
            }

            Marshal.ReleaseComObject(activeSheet);
        }

        public void OnClearViews(IRibbonControl control)
        {
            IEnumerable<IExcelTemplateView> views = ETKExcel.TemplateManager.GetActiveSheetViews();
            ETKExcel.TemplateManager.RemoveViews(views);
        }

        private void ShowCustomers()
        {
            if(mainCustomersView != null)
                ETKExcel.TemplateManager.RemoveView(mainCustomersView);

            mainCustomersView = ETKExcel.TemplateManager.AddView("TemplatesCustomers", "Main", "Customers", "B2");
            //mainCustomersView = ETKExcel.TemplateManager.AddView("TemplatesCustomers", "MainHorizontal", "Customers", "B2");

            mainCustomersView.SetDataSource(CustomersManager.Customers);
            mainCustomersView.Render();
        }

        private void ShowShares()
        {
            if (mainSharesView != null)
                ETKExcel.TemplateManager.RemoveView(mainSharesView);

            mainSharesView = ETKExcel.TemplateManager.AddView("TemplatesShares", "Main", "Shares", "B2");

            BasketManager basketManager = new BasketManager();
            mainSharesView.SetDataSource(basketManager);
            mainSharesView.Render();

            //mainSharesView.ViewSheetIsActivated += () => basketManager.StartChanging();
            mainSharesView.ViewSheetIsDeactivated += () => basketManager.StopChanging();
        }
        #endregion
    }
}
