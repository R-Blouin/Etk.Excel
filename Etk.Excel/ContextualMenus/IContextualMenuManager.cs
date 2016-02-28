namespace Etk.Excel.ContextualMenus
{
    using System.Collections.Generic;
    using Microsoft.Office.Interop.Excel;

    public delegate IEnumerable<IContextualMenu> ContextualMenusRequestedHandler(Worksheet sheet, Range range); 

    public interface IContextualMenuManager
    {
        event ContextualMenusRequestedHandler OnContextualMenusRequested;

        void RegisterMenuDefinitionsFromXml(string xml);
        IContextualMenu GetContextualMenu(string name);
    }
}
