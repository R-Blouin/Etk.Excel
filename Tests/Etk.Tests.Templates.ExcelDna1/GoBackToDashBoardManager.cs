using Etk.Excel;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Tests.Templates.ExcelDna1
{
    public class GoBackToDashBoardManager
    {
        public void GoBackToDashboard()
        {
            ExcelInterop.Workbook workbook = null;
            ExcelInterop.Worksheet dashBoard = null;
            try
            {
                workbook = ETKExcel.ExcelApplication.Application.ActiveWorkbook;
                dashBoard = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbook, "Dashboard");
                if (dashBoard != null)
                    dashBoard.Activate();
            }
            finally
            {
                if (dashBoard != null)
                {
                    Marshal.ReleaseComObject(dashBoard);
                    dashBoard = null;
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }
            }
        }
    }
}
