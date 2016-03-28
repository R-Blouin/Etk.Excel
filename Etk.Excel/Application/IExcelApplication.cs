using System;
using System.Windows.Forms;

namespace Etk.Excel.Application
{
    /// <summary>Wrapper and helpers around Excel application</summary>
    public interface IExcelApplication
    {
        /// <summary> Get the Excel application interop wrapper</summary>
        Microsoft.Office.Interop.Excel.Application Application { get; }

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

        // <summary>Display a dialog box for selecting an Excel Range</summary>
        /// <param name="title">Title of the message box. If none supplied then the title is 'Select a Range'</param>
        /// <returns>The selection concernedRange or null if no ranges selected</returns>
        Microsoft.Office.Interop.Excel.Range RangeSelectionDialog(string title);

        /// <summary> Return the Excel application active sheet</summary>
        Microsoft.Office.Interop.Excel.Worksheet GetActiveSheet();

        /// <summary> Return the sheet having 'name' as name owned by the given workbook</summary>
        Microsoft.Office.Interop.Excel.Worksheet GetWorkSheetFromName(Microsoft.Office.Interop.Excel.Workbook workbook, string name);
    }
}
