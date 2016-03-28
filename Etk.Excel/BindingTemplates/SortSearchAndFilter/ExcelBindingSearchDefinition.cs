using System;
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

            trimmedDefinition = trimmedDefinition.Replace(Search_PREFIX, string.Empty);
            trimmedDefinition = trimmedDefinition.Replace(Search_POSTFIX, string.Empty);
            return new ExcelBindingSearchDefinition(trimmedDefinition);
        }

        public override BindingSearchContextItem CreateContextItem(ITemplateView view)
        {
            return new ExcelBindingSearchContextItem(view, this);
        }
        #endregion
    }
}
