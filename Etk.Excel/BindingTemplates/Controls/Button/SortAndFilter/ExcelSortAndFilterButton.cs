using System;
using System.Reflection;
using System.Threading;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.Extensions;
using Etk.Excel.UI.Windows.SortAndFilter;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Controls.Button.SortAndFilter
{
    using ExcelForms = Microsoft.Vbe.Interop.Forms;

    class ExcelSortAndFilterButton : IDisposable
    {
        #region attributes and properties
        private static int cpt = 0;

        protected ExcelForms.CommandButton commandButton;
        protected ExcelForms.CommandButtonEvents_ClickEventHandler CurrentOnClick
        { get; private set; }

        public string Name
        { get; protected set; }

        public bool IsDisposed
        { get; protected set; }

        public ExcelTemplateView View
        { get; protected set; }

        public Range OwnerRange
        { get; protected set; }

        public ExcelForms.Font Font
        { get { return commandButton == null ? null : commandButton.Font; } }
        #endregion

        #region .ctors
        public ExcelSortAndFilterButton(ExcelTemplateView templateView)
        {
            this.View = templateView;
            Worksheet worksheet = View.SheetDestination;
            OwnerRange = View.FirstOutputCell;
            Name = string.Format("ExcelBtn{0}", Interlocked.Increment(ref cpt));

            Shape shape = (Shape)worksheet.Shapes.AddOLEObject("Forms.CommandButton.1",
                                                                Type.Missing,
                                                                false,
                                                                false,
                                                                Type.Missing,
                                                                Type.Missing,
                                                                Type.Missing,
                                                                OwnerRange.Left,
                                                                OwnerRange.Top,
                                                                20,
                                                                20);

            shape.Name = Name;
            object s = worksheet.GetType().InvokeMember(Name, BindingFlags.GetProperty, null, worksheet, null);
            commandButton = s as ExcelForms.CommandButton;
 
 
            commandButton.FontName = "Arial";
            commandButton.Font.Size = 8;
            commandButton.Caption = "S/F";
            commandButton.ForeColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            commandButton.Click += () => {  
                                            using(ExcelMainWindow excelWindow = new ExcelMainWindow(View.SheetDestination.Application.Hwnd))
                                            {
                                                SortAndFilterManagement.DisplaySortAndFilterWindow(excelWindow, View);
                                            }
                                         };

            worksheet = null;
        }
        #endregion

        public void Dispose()
        {
            if (commandButton != null)
            {
                IsDisposed = true;
 
                View.SheetDestination.OLEObjects(Name).Delete();
                commandButton = null;
                OwnerRange = null;
            }
        }
    }
}
