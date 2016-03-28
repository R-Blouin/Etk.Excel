using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public class FilterOnValues : IFilterDefinition
    {
        public ITemplateDefinition TemplateDefinition
        { get; private set; }

        public IBindingDefinition DefinitionToFilter
        { get; private set; }

        public string FilterExpression
        { get; protected set; }

        public bool UseOrEquals
        { get; private set;  }

        public IEnumerable<object> SelectedValues
        { get; set;  }

        public FilterOnValues(ITemplateDefinition templateDefinition, IBindingDefinition definition, IEnumerable<object> selectedValues, bool useOrEquals)
        {
            TemplateDefinition = templateDefinition;
            DefinitionToFilter = definition;
            SelectedValues = selectedValues;
            UseOrEquals = useOrEquals;
            if (selectedValues != null && selectedValues.Any())
                SetFilterExpression();
        }

        private void SetFilterExpression()
        {
            const string formatForOrEqual = "{0} == {1}";
            const string formatNoForOrEqual = "{0} != {1}";
            const string formatForOrEqualString = "{0} == \"{1}\"";
            const string formatNoForOrEqualString = "{0} != \"{1}\"";
            const string formatForOrEqualDateTime = "{0}.Ticks == {1}";
            const string formatNoForOrEqualDateTime = "{0}.Ticks != {1}";

            List<object> toWorkWith = new List<object>(SelectedValues);
            string[] expressionArray = new string[toWorkWith.Count];
            bool isString = DefinitionToFilter.BindingType == typeof(string);
            bool isDateTime = DefinitionToFilter.BindingType == typeof(DateTime);
            for(int i = 0; i < toWorkWith.Count; i++)
            {
                object o = toWorkWith[i];
                if (o == null)
                    o = expressionArray[i] = string.Format("{0} == null");
                else
                {
                    if (isString)
                    {
                        if (UseOrEquals)
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatForOrEqualString, DefinitionToFilter.Name, o);
                        else
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatNoForOrEqualString, DefinitionToFilter.Name, o);
                    }
                    else if (isDateTime)
                    {
                        if (UseOrEquals)
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatForOrEqualDateTime, DefinitionToFilter.Name, ((DateTime)o).Ticks);
                        else
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatNoForOrEqualDateTime, DefinitionToFilter.Name, ((DateTime)o).Ticks);
                    }
                    else
                    {
                        if (UseOrEquals)
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatForOrEqual, DefinitionToFilter.Name, o);
                        else
                            expressionArray[i] = string.Format(CultureInfo.InvariantCulture, formatNoForOrEqual, DefinitionToFilter.Name, o);
                    }
                }
            }
            FilterExpression = string.Join(UseOrEquals ? " OR " : " AND ", expressionArray);
        }
    }
}
