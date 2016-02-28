namespace Etk.Excel.Application
{
    using System;
    using System.ComponentModel.Composition;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Forms;
    using System.Windows.Threading;
    using Etk.Excel.Extensions;
    using Microsoft.Office.Core;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary> Implements <see cref="IExcelApplication"/> </summary> 
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    class ExcelApplication : IExcelApplication,  IDisposable
    {
        #region attribute and properties
        private bool isDisposed;
        private readonly object syncObj = new object();
        private CommandBarControl newMenu;
        private ExcelPostAsynchronousManager postAsynchronousManager;

        /// <summary> Implements <see cref="IExcelApplication.Application"/> </summary> 
        public Excel.Application Application
        { get; private set; }

        public Dispatcher ExcelDispatcher
        { get; private set; }
        #endregion

        #region .ctors
        [ImportingConstructor]
        public ExcelApplication([Import] Excel.Application application)
        {
            try
            {
                if (application == null)
                    throw new EtkException("the 'application' parameter is mandatory");
                Application = application;
                ExcelDispatcher = Dispatcher.CurrentDispatcher;
                newMenu = Application.CommandBars["Worksheet Menu Bar"].FindControl(1,
                                                                                    18, //the item to look for
                                                                                    Type.Missing, //the tag property (in this case missing)
                                                                                    Type.Missing, //the visible property (in this case missing)
                                                                                    true); //we want to look for it recursively
                ExcelDispatcher.ShutdownStarted += (s, o) => ETKExcel.Instance.Dispose();
                postAsynchronousManager = new ExcelPostAsynchronousManager(ExcelDispatcher);
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("ExcelApplication initialization failed:{0}", ex.Message), ex);
            }
        }

        ~ExcelApplication()
        {
            Dispose();
        }
        #endregion

        #region public methods
        /// <summary> Implements <see cref="IExcelApplication.IsInEditMode"/> </summary> 
        public bool IsInEditMode()
        {
            if (Application != null && newMenu == null)
                return false;
            return ! newMenu.Enabled;
        }

        /// <summary> Implements <see cref="IExcelApplication.DisplayException"/> </summary> 
        public void DisplayException(string title, string message, Exception ex)
        {
            StringBuilder builder = new StringBuilder(message);

            if (string.IsNullOrEmpty(title))
                title = "ETK";

            Exception currentEx = ex;
            while (currentEx != null)
            {
                builder.AppendFormat("\n\r{0}", currentEx.Message);
                currentEx = currentEx.InnerException;
            }

            if (Application != null)
            {
                using (ExcelMainWindow mainWindow = new ExcelMainWindow(Application.Hwnd))
                {
                    MessageBox.Show(mainWindow, builder.ToString(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                MessageBox.Show(builder.ToString(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary> Implements <see cref="IExcelApplication.DisplayMessageBox"/> </summary> 
        public void DisplayMessageBox(string title, string message, MessageBoxIcon icon)
        {
            if (string.IsNullOrEmpty(title))
                title = "ETK";

            if (Application != null)
            {
                using (ExcelMainWindow mainWindow = new ExcelMainWindow(Application.Hwnd))
                {
                    MessageBox.Show(mainWindow, message, title, MessageBoxButtons.OK, icon);
                }
            }
            else
                MessageBox.Show(message, title, MessageBoxButtons.OK, icon);
        }

        /// <summary> Implements <see cref="IExcelApplication.PostAsynchronousAction"/> </summary> 
        public void PostAsynchronousAction(Action action)
        {
            postAsynchronousManager.PostAction(action);
        }

        /// <summary> Implements <see cref="IExcelApplication.RangeSelectionDialog"/> </summary> 
        public Excel.Range RangeSelectionDialog(string title)
        {
            Excel.Range selectedRange = null;
            if (string.IsNullOrEmpty(title))
                title = "Select a Range";

            object obj = Application.InputBox(title, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, 8);
            if (obj is Excel.Range)
                selectedRange = obj as Excel.Range;
            return selectedRange;
        }

        /// <summary> Implements <see cref="IExcelApplication.GetActiveSheet"/> </summary> 
        public Excel.Worksheet GetActiveSheet()
        {
            Excel.Worksheet ret = null;
            if (Application != null)
                ret = Application.ActiveSheet; 
            return ret;
        }

        /// <summary> Implements <see cref="IExcelApplication.GetWorkSheetFromName"/> </summary> 
        public Excel.Worksheet GetWorkSheetFromName(Excel.Workbook workbook, string name)
        {
            if(workbook != null && ! string.IsNullOrEmpty(name))
            {
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (string.Equals(sheet.Name, name))
                        return sheet;
                }
            }
            return null;
        }

        public void Dispose()
        {
            lock (syncObj)
            {
                if (!isDisposed)
                {
                    isDisposed = true;
                    postAsynchronousManager.Dispose();
                    Marshal.ReleaseComObject(Application);
                    Application = null;
                    ExcelDispatcher = null;
                }
            }
        }
        #endregion
    }
}
