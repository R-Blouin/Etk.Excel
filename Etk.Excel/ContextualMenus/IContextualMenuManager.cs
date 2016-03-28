using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.ContextualMenus
{
    public delegate IEnumerable<IContextualMenu> ContextualMenusRequestedHandler(Worksheet sheet, Range range); 

    public interface IContextualMenuManager
    {
        event ContextualMenusRequestedHandler OnContextualMenusRequested;

        void RegisterMenuDefinitionsFromXml(string xml);
        IContextualMenu GetContextualMenu(string name);
    }
}
