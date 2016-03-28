using System.Collections.Generic;

namespace Etk.Excel.ContextualMenus
{
    public interface IContextualMenu: IContextualPart
    {
        string Name { get; }
        string Caption { get; }
        bool BeginGroup  { get; }
        int InsertBefore { get; }

        IEnumerable<IContextualPart> Items { get; }
    }
}
