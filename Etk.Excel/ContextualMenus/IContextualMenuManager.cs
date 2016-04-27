using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.ContextualMenus
{
    public delegate IEnumerable<IContextualMenu> ContextualMenusRequestedHandler(ExcelInterop.Worksheet sheet, ExcelInterop.Range range); 

    public interface IContextualMenuManager
    {
        event ContextualMenusRequestedHandler OnContextualMenusRequested;

        void RegisterMenuDefinitionsFromXml(string xml);
        IContextualMenu GetContextualMenu(string name);
    }
}
