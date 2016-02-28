﻿namespace Etk.Excel.BindingTemplates.Controls
{
    using Etk.BindingTemplates.Context;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    class EnumManager
    {
        private Dictionary<Type, string> enumByType = new Dictionary<Type, string>();

        public void CreateControl(IBindingContextItem item, ref Range range)
        {
            string values;
            if (!enumByType.TryGetValue(item.BindingDefinition.BindingType, out values))
            {
                Type type = item.BindingDefinition.IsNullable ? item.BindingDefinition.BindingType.GetGenericArguments()[0] : item.BindingDefinition.BindingType;

                List<string> list = new List<string>();
                if (item.BindingDefinition.IsNullable)
                    list.Add(string.Empty);

                list.AddRange(Enum.GetNames(type).OrderBy(s => s));

                string separator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
                values = string.Join(separator, list);
                enumByType[item.BindingDefinition.BindingType] = values;
            }

            range.Validation.Add(XlDVType.xlValidateList,
                                 XlDVAlertStyle.xlValidAlertInformation,
                                 XlFormatConditionOperator.xlBetween,
                                 values,
                                 Type.Missing);
            range.Validation.IgnoreBlank = false;
            range.Validation.InCellDropdown = true;
        }
    }
}
