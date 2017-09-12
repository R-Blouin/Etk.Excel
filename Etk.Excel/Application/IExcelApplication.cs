using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Application
{
    /// <summary>Wrapper and helpers around Excel application</summary>
    public interface IExcelApplication
    {
        /// <summary> Get the Excel application interop wrapper</summary>
        ExcelInterop.Application Application { get; }

        /// <summary> Determinate if Excel is in 'edit mode'</summary>
        /// <returns>True if Excel is in 'edit mode'</returns>
        bool IsInEditMode();
        
        /// <summary> Display an exeption message box using Excel as dialog owner</summary>
        /// <param name="title">Title of the message box</param>
        /// <param name="message">Message to display</param>
        /// <param name="ex">Exception to display</param>
        void DisplayException(string title, string message, Exception ex);

        /// <summary> Display a message Box using Excel as dialog owner</summary>
        /// <param name="title">Title of the message box</param>
        /// <param name="message">Message to display</param>
        /// <param name="icon">Icon to display in the message box</param>
        void DisplayMessageBox(string title, string message, MessageBoxIcon icon);

        /// <summary>Execute a 'System.Action' asynchronously in Excel</summary>
        /// <param name="action">System.Action to execute</param>
        void PostAsynchronousAction(System.Action action);

        /// <summary>Execute a collection of 'System.Action' synchronously (one after the other) asynchronously in Excel</summary>
        /// <param name="actions">The actions to execute</param>
        /// <param name="postExecutionAction">An action executed at the end of the execution of the 'actions'</param>
        void PostAsynchronousActions(IEnumerable<Action> actions, Action postExecutionAction);

        // <summary>Display a dialog box for selecting an Excel Range</summary>
        /// <param name="title">Title of the message box. If none supplied then the title is 'Select a Range'</param>
        /// <returns>The selection concernedRange or null if no ranges selected</returns>
        ExcelInterop.Range RangeSelectionDialog(string title);

        /// <summary> Return the Excel application active sheet</summary>
        ExcelInterop.Worksheet GetActiveSheet();

        /// <summary> Return the sheet having 'name' as name owned by the given workbook</summary>
        ExcelInterop.Worksheet GetWorkSheetFromName(ExcelInterop.Workbook workbook, string name);

        /// <summary> Indicates whether status remains visible</summary>
        bool KeepStatusVisible { get; set; }

        void ExecuteVbaMAcro(string functionName, object[] parameters);
    }
}
