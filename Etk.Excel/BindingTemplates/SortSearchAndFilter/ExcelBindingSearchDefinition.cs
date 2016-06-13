using System;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Views;

namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    public class ExcelBindingSearchDefinition : BindingSearchDefinition
    {
        #region attributes and properties
        private const string ExceptionTextFormat = "Search definition  '{0}' is invalid. The correct definition is '{1}<watermark>'.";
        public const string Search_PREFIX = "{?";
        public const string Search_POSTFIX = "?}";
        #endregion

        #region .ctors and factories
        private ExcelBindingSearchDefinition(string watermark)
                                            : base(watermark)
        {}

        public static ExcelBindingSearchDefinition CreateInstance(string trimmedDefinition)
        {
            if (!trimmedDefinition.EndsWith(Search_POSTFIX))
                throw new Exception(string.Format(ExceptionTextFormat, trimmedDefinition, Search_PREFIX, Search_POSTFIX));

            string watermark = trimmedDefinition.Replace(Search_PREFIX, string.Empty);
            watermark = watermark.Replace(Search_POSTFIX, string.Empty);
            return new ExcelBindingSearchDefinition(watermark);
        }

        public override BindingSearchContextItem CreateContextItem(ITemplateView view, IBindingContextElement parent)
        {
            return new ExcelBindingSearchContextItem(view, this, parent);
        }
        #endregion
    }
}
