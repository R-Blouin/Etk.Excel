namespace Etk.Excel.Application
{
    using System.Collections.Generic;
    using Etk.BindingTemplates.Context;
    using Etk.Excel.BindingTemplates.Views;

    class ExcelNotityPropertyContext
    {
        public IBindingContextItem ContextItem
        { get; private set; }

        public ExcelTemplateView View
        { get; private set; }

        public KeyValuePair<int, int> Param
        { get; private set; }

        public bool ChangeColor
        { get; private set; }

        public ExcelNotityPropertyContext(IBindingContextItem contextItem, ExcelTemplateView view, KeyValuePair<int, int> param, bool changeColor = false)
        {
            ContextItem = contextItem;
            View = view;
            Param = param;
            ChangeColor = changeColor;
        }
    }
}
