using System;
using System.Reflection;
using System.Threading;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using ExcelForms = Microsoft.Vbe.Interop.Forms;
using Etk.Excel.Application;

namespace Etk.Excel.BindingTemplates.Controls.CheckBox
{
    class ExcelCheckBox : IDisposable
    {
        #region attributes and properties
        private static int cpt = 0;
        private ExcelForms.CheckBox CheckBox;
        private ExcelForms.MdcCheckBoxEvents_ClickEventHandler CurrentOnClick;

        public string Name
        { get; private set; }

        public bool IsDisposed
        { get; private set; }

        public ExcelTemplateView View
        { get; private set; }

        public ExcelInterop.Range OwnerRange
        { get; private set; }

        public bool IsChecked
        {
            get { return (bool)CheckBox.get_Value(); }
            set { CheckBox.set_Value(value); }
        }
        #endregion

        #region .ctors
        public ExcelCheckBox(ExcelInterop.Range range, ExcelCheckBoxDefinition definition)
        {
            OwnerRange = range;
            OwnerRange.Value2 = null;
            ExcelInterop.Worksheet worksheet = null;
            ExcelInterop.OLEObjects oleObjects = null;
            ExcelInterop.OLEObject oleObject = null;
            try
            {
                worksheet = OwnerRange.Worksheet;
                Name = $"ExcelCB{Interlocked.Increment(ref cpt)}";

                oleObjects = worksheet.OLEObjects();
                oleObject = oleObjects.Add("Forms.CheckBox.1",
                                            Type.Missing,
                                            true,
                                            false,
                                            Type.Missing,
                                            Type.Missing,
                                            Type.Missing,
                                            OwnerRange.Left + 3,
                                            OwnerRange.Top + 1,
                                            12,
                                            12);
                oleObject.Name = Name;
                oleObject.Placement = ExcelInterop.XlPlacement.xlMove;
                CheckBox = worksheet.GetType().InvokeMember(Name, BindingFlags.Default | BindingFlags.GetProperty, null, worksheet, null) as ExcelForms.CheckBox;

                CheckBox.SpecialEffect = ExcelForms.fmButtonEffect.fmButtonEffectSunken;
                CheckBox.TripleState = false;

                CheckBox.Caption = string.Empty;
                CheckBox.BackColor = (int)OwnerRange.Interior.Color;
                CheckBox.BackStyle = ExcelForms.fmBackStyle.fmBackStyleTransparent;
                oleObject.Interior.ColorIndex = -4142;
                CheckBox.AutoSize = false;

                oleObject = null;
            }
            finally
            {
                if (oleObject != null)
                    ExcelApplication.ReleaseComObject(oleObject);
                if (oleObjects != null)
                    ExcelApplication.ReleaseComObject(oleObjects);
                if (worksheet != null)
                    ExcelApplication.ReleaseComObject(worksheet);
            }
        }

        public void SetOnClick(Action action)
        {
            if (CurrentOnClick != null)
                CheckBox.Click -= CurrentOnClick;

            if (action != null)
            {
                CurrentOnClick = () => action();
                CheckBox.Click += CurrentOnClick;
            }
        }
        #endregion

        #region public methods
        public void Dispose()
        {
            if (CheckBox != null)
            {
                IsDisposed = true;
                if (CurrentOnClick != null)
                    CheckBox.Click -= CurrentOnClick;

                ExcelInterop.Worksheet worksheet = null;
                ExcelInterop.OLEObject oleObject = null;
                try
                {
                    worksheet = OwnerRange.Worksheet;
                    oleObject = worksheet.OLEObjects(Name);
                    oleObject.Delete();
                }
                finally
                {
                    if (worksheet != null)
                        ExcelApplication.ReleaseComObject(worksheet);
                    if (oleObject != null)
                        ExcelApplication.ReleaseComObject(oleObject);
                }
                ExcelApplication.ReleaseComObject(OwnerRange);
                ExcelApplication.ReleaseComObject(CheckBox);
            }
        }
        #endregion
    }
}
