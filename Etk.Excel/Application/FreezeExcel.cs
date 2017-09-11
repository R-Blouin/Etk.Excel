using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Application
{
    /// <summary>
    /// Use to freeze the Excel execution the time to execute operations.
    /// Reduce the flickering during multiple updates on cells in the current Excel application.
    /// </summary>
    public class FreezeExcel : IDisposable
    {
        private static int requestsCpt = 0;
        private static readonly object objSync = new object();

        private bool disposed;
        private readonly bool screenUpdating;
        private readonly bool enableEvents;
        private readonly bool displayStatusBar;
        private readonly ExcelInterop.XlCalculation calculationMode;

        public FreezeExcel(bool keepStatusVisible = true, bool keepScreenUpdating = false, bool keepEnabledEvent = false, bool keepCalculation = false)
        {
            lock (objSync)
            {
                requestsCpt++;
                if (! ETKExcel.ExcelApplication.IsInEditMode())
                {
                    if (requestsCpt == 1)
                    {
                        screenUpdating = ETKExcel.ExcelApplication.Application.ScreenUpdating;
                        enableEvents = ETKExcel.ExcelApplication.Application.EnableEvents;
                        displayStatusBar = ETKExcel.ExcelApplication.Application.DisplayStatusBar;
                        calculationMode = ETKExcel.ExcelApplication.Application.Calculation;

                        ETKExcel.ExcelApplication.Application.ScreenUpdating   = keepScreenUpdating && screenUpdating;
                        ETKExcel.ExcelApplication.Application.EnableEvents     = keepEnabledEvent && enableEvents;
                        ETKExcel.ExcelApplication.Application.DisplayStatusBar = keepStatusVisible && displayStatusBar;
                        ETKExcel.ExcelApplication.Application.Calculation      = keepCalculation ? ETKExcel.ExcelApplication.Application.Calculation : ExcelInterop.XlCalculation.xlCalculationManual;
                    }
                }
            }
        }

        ~FreezeExcel()
        {
            Dispose();
        }

        public void Dispose()
        {
            if (!disposed)
            {
                disposed = true;
                lock (objSync)
                {
                    requestsCpt--;
                    if (! ETKExcel.ExcelApplication.IsInEditMode())
                    {
                        if (requestsCpt == 0)
                        {
                            ETKExcel.ExcelApplication.Application.ScreenUpdating = screenUpdating;
                            ETKExcel.ExcelApplication.Application.EnableEvents = enableEvents;
                            ETKExcel.ExcelApplication.Application.DisplayStatusBar = displayStatusBar;
                            ETKExcel.ExcelApplication.Application.Calculation = calculationMode;
                        }
                    }
                }
            }
        }
    }
}
