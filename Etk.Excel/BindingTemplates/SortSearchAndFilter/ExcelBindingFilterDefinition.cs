namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
    using Etk.BindingTemplates.Views;
    using Etk.Excel.BindingTemplates.Definitions;

    public class ExcelBindingFilterDefinition : BindingFilterDefinition
    {
        #region attributes and properties
        private const string ExceptionTextFormat = "Filter definition  '{0}' is invalid. The correct definition is '{1}<watermark>:property to filter path{2}'.";
        public const string Filter_PREFIX = "{*";
        public const string Filter_POSTFIX = "*}";
        #endregion

        #region .ctors and factories
        private ExcelBindingFilterDefinition(ExcelTemplateDefinitionPart templateDefinitionPart, string definition, string watermark, IEnumerable<string> path)
                                            : base(templateDefinitionPart, definition, watermark, path)
        {}

        public static ExcelBindingFilterDefinition CreateInstance(ExcelTemplateDefinitionPart templateDefinitionPart, string trimmedDefinition)
        {
            if (!trimmedDefinition.EndsWith(Filter_POSTFIX))
                throw new Exception(string.Format(ExceptionTextFormat, trimmedDefinition, Filter_PREFIX, Filter_POSTFIX));

            trimmedDefinition = trimmedDefinition.Replace(Filter_PREFIX, string.Empty);
            trimmedDefinition = trimmedDefinition.Replace(Filter_POSTFIX, string.Empty);
            if(string.IsNullOrEmpty(trimmedDefinition))
                throw new Exception(string.Format(ExceptionTextFormat, trimmedDefinition, Filter_PREFIX, Filter_POSTFIX));

            string[] defParts = trimmedDefinition.Split(':');

            string watermark = null;
            string definitionPath;
            switch (defParts.Count())
            { 
                case 1:
                    definitionPath = defParts[0];
                break;
                case 2:
                    watermark = defParts[0];
                    definitionPath = defParts[1];
                break;
                default:
                    throw new Exception(string.Format(ExceptionTextFormat, trimmedDefinition, Filter_PREFIX, Filter_POSTFIX));
            }

            if(string.IsNullOrEmpty(definitionPath))
                throw new Exception(string.Format(ExceptionTextFormat, trimmedDefinition, Filter_PREFIX, Filter_POSTFIX));

            string[] path = definitionPath.Split('-');
            return new ExcelBindingFilterDefinition(templateDefinitionPart, trimmedDefinition, watermark, path);
        }

        override public BindingFilterContextItem CreateContextItem(ITemplateView view, IBindingContextElement contextElement)
        {
            return new ExcelBindingFilterContextItem(view, this, contextElement);
        }
        #endregion
    }
}
