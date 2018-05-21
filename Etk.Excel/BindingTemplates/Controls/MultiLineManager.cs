using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
                {
                    ExcelInterop.Range mergeArea = range.MergeArea;
                    ExcelInterop.Range columns = mergeArea.Columns;
                    hOffset = columns.Count;
                    Marshal.ReleaseComObject(mergeArea);
                    Marshal.ReleaseComObject(columns);
                }

                IEnumerable<BorderStyle> bordersStyle = RetrieveBorders(source.MergeCells ? source.MergeArea : source);
                ExcelInterop.Range toMerge = range.Resize[vOffset, hOffset];
                toMerge.Merge();

                ExcelInterop.Borders borders = toMerge.Borders;
                foreach (BorderStyle borderStyle in bordersStyle)
                {
                    ExcelInterop.Border border = borders[borderStyle.Index];

                    border.ColorIndex = borderStyle.ColorIndex;
                    border.Weight = borderStyle.Weight;
                    border.LineStyle = borderStyle.LineStyle;
                    border.Color = borderStyle.Color;
                    border.TintAndShade = borderStyle.TintAndShade;

                    Marshal.ReleaseComObject(border);
                }
                Marshal.ReleaseComObject(borders);
                Marshal.ReleaseComObject(toMerge);
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(source);
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
            ExcelInterop.Borders borders = range.Borders;
            ExcelInterop.Border border = borders[bordersIndex];
            if (border.LineStyle != (int) ExcelInterop.XlLineStyle.xlLineStyleNone)
            {
                ret = new BorderStyle();
                ret.Index = bordersIndex;
                ret.LineStyle = border.LineStyle;
                ret.Weight = border.Weight;
                ret.ColorIndex = border.ColorIndex;
                ret.Color = border.Color;
                ret.TintAndShade = border.TintAndShade;
            }
            Marshal.ReleaseComObject(border);
            Marshal.ReleaseComObject(borders);
            return ret;
        }
    }
}
