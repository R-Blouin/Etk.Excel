using System.Linq;
using Etk.BindingTemplates.Context;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Controls
{
    class MultiLineManager
    {
        public void CreateControl(IBindingContextItem item, ref Range range, ref int vOffset)
        {
            int hOffset = 1;
            if (item.BindingDefinition.MultiLineFactorResolver != null)
            {
                object toInvoke = item.BindingDefinition.MultiLineFactorResolver.IsStatic ? null : item.ParentElement.DataSource;
                object[] parameters = item.BindingDefinition.MultiLineFactorResolver.GetParameters().Length == 0 ? null : new object[] { item.ParentElement.DataSource };
                vOffset = (int) item.BindingDefinition.MultiLineFactorResolver.Invoke(toInvoke, parameters);
                if (vOffset <= 0)
                    vOffset = 1;
            }
            else
            {
                object objValue = item.ResolveBinding();
                if (objValue is string)
                {
                    string value = objValue as string;
                    int nbrLine = value.Count(c => c.Equals('\n'));
                    if (nbrLine > 0)
                    {
                        vOffset = (int)((nbrLine + 1) * item.BindingDefinition.MultiLineFactor);
                        if (range.MergeCells)
                            hOffset = range.Columns.Count;
                    }
                }
            }
        
            Range toMerge = range.Resize[vOffset, hOffset];
            toMerge.Merge();
        }
    }
}
