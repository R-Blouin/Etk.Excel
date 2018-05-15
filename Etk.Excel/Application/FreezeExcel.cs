using System;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Application
{
    /// <summary>
    /// Use to freeze the Excel execution the time to execute operations.
    /// Reduce the flickering during multiple updates on cells in the current Excel application.
    /// </summary>
    public class FreezeExcel : IDisposable
    {
        private static int requestsCpt;
        private static readonly object objSync = new object();

        private bool disposed;
        private readonly bool screenUpdating;
        private readonly bool enableEvents;
        private readonly bool displayStatusBar;
        private readonly ExcelInterop.XlCalculation calculationMode;

        #region .ctors
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

                        Freeze(keepStatusVisible, keepScreenUpdating, keepEnabledEvent, keepCalculation);
                    }
                }
            }
        }

        ~FreezeExcel()
        {
            Dispose();
        }
        #endregion

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
                            UnFreeze();
                    }
                }
            }
        }

        private void Freeze(bool keepStatusVisible, bool keepScreenUpdating, bool keepEnabledEvent, bool keepCalculation)
        {
            try
            {
                ETKExcel.ExcelApplication.Application.ScreenUpdating = keepScreenUpdating && screenUpdating;
                ETKExcel.ExcelApplication.Application.EnableEvents = keepEnabledEvent && enableEvents;
                ETKExcel.ExcelApplication.Application.DisplayStatusBar = keepStatusVisible && displayStatusBar;
                ETKExcel.ExcelApplication.Application.Calculation = keepCalculation ? ETKExcel.ExcelApplication.Application.Calculation : ExcelInterop.XlCalculation.xlCalculationManual;
            }
            catch (COMException comEx)
            {
                if (comEx.ErrorCode == ETKExcel.EXCEL_BUSY)
                {
                    Thread.Sleep(ETKExcel.WAITINGTIME_EXCEL_BUSY);
                    Freeze(keepStatusVisible, keepScreenUpdating, keepEnabledEvent, keepCalculation);
                    return;
                }

                throw new EtkException($"'Freeze Excel' failed: {comEx.Message}");
            }
        }

        private void UnFreeze()
        {
            try
            {
                ETKExcel.ExcelApplication.Application.ScreenUpdating = screenUpdating;
                ETKExcel.ExcelApplication.Application.EnableEvents = enableEvents;
                ETKExcel.ExcelApplication.Application.DisplayStatusBar = displayStatusBar;
                ETKExcel.ExcelApplication.Application.Calculation = calculationMode;
            }
            catch (COMException comEx)
            {
                if (comEx.ErrorCode == ETKExcel.EXCEL_BUSY)
                {
                    Thread.Sleep(ETKExcel.WAITINGTIME_EXCEL_BUSY);
                    UnFreeze();
                    return;
                }

                throw new EtkException($"'UnFreeze Excel' failed: {comEx.Message}");
            }
        }
    }
}
