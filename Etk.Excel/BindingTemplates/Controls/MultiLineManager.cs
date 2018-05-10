using System;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Controls
{
    class BorderStyle
    {
        public dynamic Index;
        public dynamic LineStyle;
        public dynamic Weight;
        public dynamic ColorIndex;
        public dynamic Color;
        public dynamic TintAndShade;
    }

    class MultiLineManager
    {
        public void CreateControl(IBindingContextItem item, ExcelInterop.Range range, ExcelInterop.Range source, ref int vOffset)
        {
            try
            {
                int hOffset = 1;
                if (item.BindingDefinition.MultiLineFactorResolver != null)
                {
                    if (!item.BindingDefinition.MultiLineFactorResolver.IsNotDotNet)
                    {
                        object toInvoke = item.BindingDefinition.MultiLineFactorResolver.Callback.IsStatic ? null : item.ParentElement.DataSource;
                        if (toInvoke != null)
                        {
                            object[] parameters = item.BindingDefinition.MultiLineFactorResolver.Callback.GetParameters().Length == 0 ? null : new [] {item.ParentElement.DataSource};
                            vOffset = (int)  item.BindingDefinition.MultiLineFactorResolver.Callback.Invoke(toInvoke, parameters);
                            if (vOffset <= 0)
                                vOffset = 1;
                        }
                    }
                }
                else
                {
                    object objValue = item.ResolveBinding();
                    if (objValue is string)
                    {
                        string value = objValue as string;
                        int nbrLine = value.Count(c => c.Equals('\n'));
                        if (nbrLine > 0)
                            vOffset = (int) ((nbrLine + 1)*item.BindingDefinition.MultiLineFactor);
                    }
                }

                if (range.MergeCells)
                    hOffset = range.MergeArea.Columns.Count;

                IEnumerable<BorderStyle> bordersStyle = RetrieveBorders(source.MergeCells ? source.MergeArea : source);
                ExcelInterop.Range toMerge = range.Resize[vOffset, hOffset];
                toMerge.Merge();

                foreach (BorderStyle borderStyle in bordersStyle)
                {
                    toMerge.Borders[borderStyle.Index].ColorIndex = borderStyle.ColorIndex;
                    toMerge.Borders[borderStyle.Index].Weight = borderStyle.Weight;
                    toMerge.Borders[borderStyle.Index].LineStyle = borderStyle.LineStyle;
                    toMerge.Borders[borderStyle.Index].Color = borderStyle.Color;
                    toMerge.Borders[borderStyle.Index].TintAndShade = borderStyle.TintAndShade;
                }
            }
            catch (Exception ex)
            {
                throw new EtkException($"MultiLine manager failed (option 'ME=' of the cell '{item.BindingDefinition.BindingExpression}': {ex.Message}");
            }
        }

        private IEnumerable<BorderStyle> RetrieveBorders(ExcelInterop.Range range)
        {
            List<BorderStyle> ret = new List<BorderStyle>();
            foreach (ExcelInterop.XlBordersIndex styleIndex in Enum.GetValues(typeof(ExcelInterop.XlBordersIndex)))
            {
                BorderStyle style = RetrieveBorderStyle(range, styleIndex);
                if(style != null)
                    ret.Add(style);
            }
            return ret;
        }

        private BorderStyle RetrieveBorderStyle(ExcelInterop.Range range, ExcelInterop.XlBordersIndex bordersIndex)
        {
            BorderStyle ret = null;
            if (range.Borders[bordersIndex].LineStyle != (int)ExcelInterop.XlLineStyle.xlLineStyleNone)
            {
                ret = new BorderStyle();
                ret.Index = bordersIndex;
                ret.LineStyle = range.Borders[bordersIndex].LineStyle;
                ret.Weight =  range.Borders[bordersIndex].Weight;
                ret.ColorIndex = range.Borders[bordersIndex].ColorIndex;
                ret.Color = range.Borders[bordersIndex].Color;
                ret.TintAndShade = range.Borders[bordersIndex].TintAndShade;
            }
            return ret;
        }
    }
}
