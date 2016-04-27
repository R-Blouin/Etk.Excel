using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Controls.Button
{
    using ExcelForms = Microsoft.Vbe.Interop.Forms;

    class ExcelButton : IDisposable
    {
        #region attributes and properties
        private static int cpt = 0;
        
        protected ExcelForms.CommandButton commandButton;
        protected ExcelForms.CommandButtonEvents_ClickEventHandler CurrentCommand 
        { get; private set; }

        public string Name
        { get; protected set; }

        public bool IsDisposed
        { get; protected set; }

        public ExcelTemplateView View
        { get; protected set; }

        public ExcelInterop.Range OwnerRange
        { get; protected set; }

        public ExcelForms.Font Font
        { get { return commandButton == null ? null : commandButton.Font; } }

        public ExcelInterop.XlPlacement Placement
        {
            get
            {
                if (! IsDisposed)
                    return (commandButton as ExcelInterop.Shape).Placement;
                return ExcelInterop.XlPlacement.xlFreeFloating;
            }
            set 
            {
                if (! IsDisposed)
                    (commandButton as ExcelInterop.Shape).Placement = value;
            }
        }

        public String Text
        {
            get { return IsDisposed ? null : commandButton.Caption; }
            set 
            {
                if (! IsDisposed)
                    commandButton.Caption = value; 
            }
        }

        public bool Enable
        {
            get { return IsDisposed ? false : commandButton.Enabled; }
            set
            {
                if (!IsDisposed)
                    commandButton.Enabled = value;
            }
        }
        #endregion

        #region .ctors
        public ExcelButton(ExcelInterop.Range range, ExcelButtonDefinition definition)
        {
            OwnerRange = range;
            OwnerRange.Value2 = null;
            ExcelInterop.Worksheet worksheet = OwnerRange.Worksheet;
            Name = string.Format("ExcelBtn{0}", Interlocked.Increment(ref cpt));

            ExcelInterop.OLEObjects oleObjects = worksheet.OLEObjects();

            ExcelInterop.OLEObject obj = oleObjects.Add("Forms.CommandButton.1", 
                                           Type.Missing,
                                           false,
                                           false,
                                           Type.Missing,
                                           Type.Missing,
                                           Type.Missing,
                                           OwnerRange.Left + definition.X,
                                           OwnerRange.Top + definition.Y,
                                           definition.W == 0 ? OwnerRange.Width : definition.W,
                                           definition.H == 0 ? OwnerRange.Height : definition.H);

            obj.Name = Name;
            object s = worksheet.GetType().InvokeMember(Name, BindingFlags.GetProperty, null, worksheet, null);
            commandButton = s as ExcelForms.CommandButton;
            commandButton.FontName = "Arial";
            commandButton.Font.Size = 8;
            commandButton.Caption = definition.Label;
            //if (excelTemplateDefinition.W == 0 && excelTemplateDefinition.H == 0)
            //    commandButton.AutoSize = true;
            obj.Placement = ExcelInterop.XlPlacement.xlMove;

            Marshal.ReleaseComObject(worksheet);
            worksheet = null;
        }
        #endregion

        #region public methods
        public void Dispose()
        {
            if (commandButton != null)
            {
                IsDisposed = true;

                if (CurrentCommand != null)
                    commandButton.Click -= CurrentCommand;

                ExcelInterop.Worksheet worksheet = OwnerRange.Worksheet;
                worksheet.OLEObjects(Name).Delete();
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
                commandButton = null;
                OwnerRange = null;
            }
        }

        public void SetCommand(MethodInfo methodInfo, object obj, bool useRange)
        {
            if (CurrentCommand != null)
                commandButton.Click -= CurrentCommand;

            if (methodInfo != null)
            {
                if (methodInfo.IsStatic)
                {
                    CurrentCommand = () => {
                                                if(Enable)
                                                {  
                                                    if(useRange)
                                                       methodInfo.Invoke(null, new object[] {obj, OwnerRange });
                                                    else
                                                       methodInfo.Invoke(null, new object[] {obj} ); 
                                                }
                                            };
                }
                else 
                {
                    CurrentCommand = () => { 
                                                if(Enable)
                                                {
                                                    if(useRange)
                                                        methodInfo.Invoke(obj, new object[] { OwnerRange });
                                                    else
                                                        methodInfo.Invoke(obj, null); 
                                                }
                                           };
                }
                commandButton.Click += CurrentCommand;
            }
        }
        #endregion
    }
}
