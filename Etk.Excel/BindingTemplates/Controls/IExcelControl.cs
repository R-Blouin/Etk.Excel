namespace Etk.Excel.BindingTemplates.Controls
{
    using Microsoft.Office.Interop.Excel;

    interface IExcelControl
    {
        void CreateControl(Range range);
    }
}
