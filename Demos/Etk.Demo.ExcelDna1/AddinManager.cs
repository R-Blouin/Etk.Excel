using Etk.Excel;
using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Demo.ExcelDna1
{
    class AddinManager : IExcelAddIn
    {
        public ExcelInterop.Application ExcelApplication
        { get; private set; }

        public void AutoOpen()
        {
            ExcelApplication = ExcelDnaUtil.Application as ExcelInterop.Application;

            // To avoid the Excel 'Save message' on Exit
            ExcelInterop.Workbook currentWorkbook = ExcelApplication.ActiveWorkbook;
            currentWorkbook.BeforeClose += (ref bool cancel) =>
            {
                //int ii = System.Runtime.InteropServices.Marshal.ReleaseComObject(ETKExcel.ExcelApplication.Application);
                currentWorkbook.Saved = true;
            };

            // Init the ETK Framework : mandatory before any uses of the framework
            ETKExcel.Init(ExcelApplication);
        }
        public void AutoClose()
        { }
    }
}
