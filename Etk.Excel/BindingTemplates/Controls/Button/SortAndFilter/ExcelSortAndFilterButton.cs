using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.Application;

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

        public ExcelInterop.Range OwnerRange
        { get; protected set; }

        public ExcelForms.Font Font => commandButton == null ? null : commandButton.Font;

        #endregion

        #region .ctors
        public ExcelSortAndFilterButton(ExcelTemplateView templateView)
        {
            this.View = templateView;
            ExcelInterop.Worksheet worksheet = null;
            ExcelInterop.Shapes shapes = null;
            ExcelInterop.Shape shape = null;
            try
            {
                worksheet = View.ViewSheet;
                OwnerRange = View.FirstOutputCell;
                Name = $"ExcelBtn{Interlocked.Increment(ref cpt)}";
                shapes = worksheet.Shapes;

                shape = (ExcelInterop.Shape) shapes.AddOLEObject("Forms.CommandButton.1",
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

                commandButton.Click += () =>
                {
                    using (ExcelMainWindow excelWindow = new ExcelMainWindow(View.ViewSheet.Application.Hwnd))
                    {
                        //@@SortAndFilterManagement.DisplaySortAndFilterWindow(excelWindow, View);
                    }
                };
            }
            finally
            {
                if (shape != null)
                    ExcelApplication.ReleaseComObject(shape);
                if (shapes != null)
                    ExcelApplication.ReleaseComObject(shapes);
                if (worksheet != null)
                    ExcelApplication.ReleaseComObject(worksheet);

                shape = null;
                shapes = null;
                worksheet = null;
            }
        }
        #endregion

        public void Dispose()
        {
            if (commandButton != null)
            {
                IsDisposed = true;
 
                View.ViewSheet.OLEObjects(Name).Delete();
                commandButton = null;
                OwnerRange = null;
            }
        }
    }
}
