using System.Linq;
using Etk.BindingTemplates.Context;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Controls
{
    class MultiLineManager
    {
        public void CreateControl(IBindingContextItem item, ref Range range, ref int vOffset)
        {
            object objValue = item.ResolveBinding();
            if (objValue is string)
            {
                string value = objValue as string;
                int nbrLine = value.Count(c => c.Equals('\n'));
                if (nbrLine > 0)
                {
                    vOffset = (int)((nbrLine + 1) * item.BindingDefinition.MultiLineFactor);
                    int hOffset = 1;
                    if (range.MergeCells)
                        hOffset = range.Columns.Count; 
                    Range toMerge = range.Resize[vOffset, hOffset];
                    toMerge.Merge();
                }
            }
        }
    }
}
