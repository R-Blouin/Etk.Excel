using System.Collections.Generic;
using Etk.BindingTemplates.Context;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.ContextualMenus
{
    class ContextualMenu : IContextualMenu 
    {
        #region properties and attributes
        public string Name
        { get; private set; }

        public string Caption
        { get; private set; }

        public bool BeginGroup
        { get; private set; }

        public int InsertBefore
        { get; private set; }

        public IEnumerable<IContextualPart> Items
        { get; private set; }
        #endregion

        #region .ctors and factories
        public ContextualMenu(string name, string caption, bool beginGroup, IEnumerable<IContextualPart> items)
        {
            Name = name;
            Caption = string.IsNullOrEmpty(caption) ? Name : caption;
            BeginGroup = beginGroup;
            Items = items;
        }
        #endregion

        #region public methods
        public void SetAction(ExcelInterop.Range range, IBindingContextElement currentContextElement, IBindingContextElement targetedContextElement)
        {
            if (Items != null)
            {
                foreach(IContextualPart part in Items)
                {
                    if (part is ContextualMenu)
                        (part as ContextualMenu).SetAction(range, targetedContextElement, currentContextElement);
                    else
                        (part as ContextualMenuItem).SetAction(range, targetedContextElement, currentContextElement);
                }
            }
        }

        public void SetAction(ExcelInterop.Range range)
        {
            if (Items != null)
            {
                foreach (IContextualPart part in Items)
                {
                    if (part is ContextualMenu)
                        (part as ContextualMenu).SetAction(range);
                    else
                        (part as ContextualMenuItem).SetAction(range);
                }
            }
        }

        //public void SetAction()
        //{
        //    if (Items != null)
        //    {
        //        foreach (IContextualPart part in Items)
        //        {
        //            if (part is ContextualMenu)
        //                (part as ContextualMenu).SetAction();
        //            else
        //                (part as ContextualMenuItem).SetAction();
        //        }
        //    }
        //}
        #endregion
    }
}
