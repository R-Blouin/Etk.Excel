using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Controls
{
    interface IExcelControl
    {
        void CreateControl(ExcelInterop.Range range);
    }
}
