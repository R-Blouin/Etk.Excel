using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Etk.Demo.Shops.UI.Excel.Sheets;
using ExcelDna.Integration.CustomUI;

namespace Etk.Demo.Shops.UI.Excel
{
    /// <summary>Manage the application custom ribbons</summary>
    [ComVisible(true)]
    public class AddinRibbon : ExcelRibbon
    {
        /// <summary>Excel Dna method used to build the ribbon</summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Demo.Shops.UI.Excel.AddinRibbon.xml")))
            {
                string ribbonXml = textReader.ReadToEnd();
                return ribbonXml;
            }
        }

        #region event Handlers
        public void OnReloadViews(IRibbonControl control)
        {
            SheetShops.Instance.RenderViews();
        }
        #endregion
    }
}
