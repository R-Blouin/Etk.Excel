using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Threading;
using Etk.Excel.Extensions;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Reflection;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Application
{
    /// <summary> Implements <see cref="IExcelApplication"/> </summary> 
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    class ExcelApplication : IExcelApplication, IDisposable
    {
        #region attribute and properties
        private bool isDisposed;
        private readonly object syncObj = new object();
        private readonly CommandBarControl newMenu;
        private readonly ExcelPostAsynchronousManager postAsynchronousManager;

        /// <summary> Implements <see cref="IExcelApplication.Application"/> </summary> 
        public ExcelInterop.Application Application
        { get; private set; }

        public Dispatcher ExcelDispatcher
        { get; private set; }

        /// <summary>
        /// Indicates whether status remains visible.
        /// </summary>
        public bool KeepStatusVisible { get; set; }
        #endregion

        #region .ctors
        [ImportingConstructor]
        public ExcelApplication([Import] ExcelInterop.Application application)
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
                //@@ExcelDispatcher.ShutdownStarted += (s, o) => ETKExcel.Instance.Dispose();
                postAsynchronousManager = new ExcelPostAsynchronousManager(ExcelDispatcher);
            }
            catch (Exception ex)
            {
                throw new EtkException($"ExcelApplication initialization failed:{ex.Message}");
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
            return !newMenu.Enabled;
        }

        /// <summary> Implements <see cref="IExcelApplication.DisplayException"/> </summary> 
        public void DisplayException(string title, string message, Exception ex)
        {
            StringBuilder builder = new StringBuilder(message);

            if (string.IsNullOrEmpty(title))
                title = "Etk";

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

        /// <summary> Implements <see cref="IExcelApplication.PostAsynchronousActions"/> </summary> 
        public void PostAsynchronousActions(IEnumerable<Action> actions, Action postExecutionAction)
        {
            if (postExecutionAction == null)
                postAsynchronousManager.PostActions(actions);
            else
            {
                ExcelPostListAsynchronousManager asynchronousManager = new ExcelPostListAsynchronousManager(ExcelDispatcher, actions, postExecutionAction);
                asynchronousManager.Execute();
            }
        }

        /// <summary> Implements <see cref="IExcelApplication.RangeSelectionDialog"/> </summary> 
        public ExcelInterop.Range RangeSelectionDialog(string title)
        {
            ExcelInterop.Range selectedRange = null;
            if (string.IsNullOrEmpty(title))
                title = "Select a Range";

            object obj = Application.InputBox(title, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            if (obj is ExcelInterop.Range)
                selectedRange = obj as ExcelInterop.Range;
            return selectedRange;
        }

        /// <summary> Implements <see cref="IExcelApplication.GetActiveSheet"/> </summary> 
        public ExcelInterop.Worksheet GetActiveSheet()
        {
            ExcelInterop.Worksheet ret = null;
            if (Application != null)
                ret = Application.ActiveSheet;
            return ret;
        }

        /// <summary> Implements <see cref="IExcelApplication.GetWorkSheetFromName"/> </summary> 
        public ExcelInterop.Worksheet GetWorkSheetFromName(ExcelInterop.Workbook workbook, string name)
        {
            if (workbook != null && !string.IsNullOrEmpty(name))
            {
                foreach (ExcelInterop.Worksheet sheet in workbook.Worksheets)
                {
                    if (string.Equals(sheet.Name, name))
                        return sheet;
                }
            }
            return null;
        }

        public void ShowHideColumns(ExcelInterop.Range targetedRange, int numberOfColumns)
        {
            StaticShowHideColumns(targetedRange, numberOfColumns);
        }

        public object ExecuteVbaMAcro(string functionName, object[] parameters)
        {
            try
            {
                object[] p;
                if (parameters == null)
                    p = new object[] { functionName };
                else
                {
                    List<object> lp = new List<object>(new object[] { functionName });
                    lp.AddRange(parameters);
                    p = lp.ToArray();
                }
                return Application.GetType().InvokeMember("Run", BindingFlags.Default | BindingFlags.InvokeMethod, null, Application, p);
            }
            catch (Exception ex)
            {
                ETKExcel.ExcelApplication.DisplayException(null, $"'Execute macro '{functionName ?? string.Empty}' failed", ex);
                return null;
            }
        }

        public void ClearRange(ExcelInterop.Range from, ExcelInterop.Range to, ExcelInterop.Range with)
        {
            if (from == null)
                return;

            ExcelInterop.Worksheet concernedSheet = null;
            bool isProtected = false;
            try
            {
                concernedSheet = from.Worksheet;

                isProtected = concernedSheet.ProtectContents;
                if (isProtected)
                    concernedSheet.Unprotect(Type.Missing);

                if (to == null)
                    to = concernedSheet.UsedRange;

                from = from.Resize[to.Rows.Count - from.Rows.Count - 1, to.Columns.Count - from.Columns.Count - 1];
                from.Clear();

                if(with != null)
                {
                    ExcelInterop.Interior withInterior = with.Interior;
                    ExcelInterop.Font withFont = with.Font;

                    ExcelInterop.Interior interior = from.Interior;
                    ExcelInterop.Font font = from.Font;

                    font.Color = withFont.Color;
                    interior.Color = withInterior.Color;

                    ExcelApplication.ReleaseComObject(interior);
                    ExcelApplication.ReleaseComObject(font);
                    ExcelApplication.ReleaseComObject(withInterior);
                    ExcelApplication.ReleaseComObject(withFont);
                    interior = null;
                    font = null;
                    withInterior = null;
                    withFont = null;
                }
            }
            catch
            {
                if(concernedSheet != null)
                    ExcelApplication.ReleaseComObject(concernedSheet);
            }
            finally
            {
                if (concernedSheet != null && isProtected)
                    ProtectSheet(concernedSheet);
            }
        }

        public void ProtectSheet(ExcelInterop.Worksheet concernedSheet)
        {
            if (concernedSheet != null && !concernedSheet.ProtectContents)
            {
                concernedSheet.Cells.Locked = false;
                concernedSheet.Protect(Type.Missing, false, false, Type.Missing, false, true,
                                       true, true,
                                       false, false,
                                       false,
                                       false, false, false, true,
                                       true);
            }
        }

        public void Dispose()
        {
            lock (syncObj)
            {
                if (!isDisposed)
                {
                    isDisposed = true;
                    postAsynchronousManager.Dispose();
                    ExcelApplication.ReleaseComObject(Application);
                    Application = null;
                    ExcelDispatcher = null;
                }
            }
        }
        #endregion

        #region static public methods
        public static void ReleaseComObject(object obj)
        {
            int refCpt = Marshal.ReleaseComObject(obj);
            if (refCpt < 0)
                MessageBox.Show("Aie !", "ReleaseComObject", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }


        public static void StaticShowHideColumns(ExcelInterop.Range targetedRange, int numberOfColumns)
        {
            if (targetedRange != null && numberOfColumns != 0)
            {
                ExcelInterop.Range workingRange;
                if (numberOfColumns < 0)
                    workingRange = targetedRange.Offset[Type.Missing, numberOfColumns];
                else
                    workingRange = targetedRange.Offset[Type.Missing, 1];

                workingRange = workingRange.Resize[Type.Missing, Math.Abs(numberOfColumns)];

                ExcelInterop.Range columns = workingRange.EntireColumn;
                columns.Hidden = !(bool)columns.Hidden;

                columns = null;
                workingRange = null;
            }
        }
        #endregion
    }
}
