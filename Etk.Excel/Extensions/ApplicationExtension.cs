using System;
using System.Windows.Forms;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.Extensions
{
    /// <summary> 
    /// Excel main window wrapper.
    /// </summary>
    public class ExcelMainWindow : NativeWindow, IDisposable
    {
        public bool IsDisposed
        { get; private set; }

        internal ExcelMainWindow(int hwnd)
        {
            IsDisposed = false;
            AssignHandle(new IntPtr(hwnd));
        }

        ~ExcelMainWindow()
        {
            Dispose();
        }

        public void Dispose()
        {
            if (! IsDisposed && ! Handle.Equals(IntPtr.Zero))
            {
                IsDisposed = true;
                ReleaseHandle();
            }
        }
    }

    /// <summary> 
    /// Extension methods for 'Microsoft.Office.Interop.Excel.Application'
    /// </summary>
    public static class ApplicationExtension
    {
        /// <summary> Return a wrapper around Excel main windows of the current application insatnce.</summary>
        public static ExcelMainWindow GetMainWindow(this ExcelInterop.Application current)
        {
            return new ExcelMainWindow(current.Hwnd);
        }
    }
}
