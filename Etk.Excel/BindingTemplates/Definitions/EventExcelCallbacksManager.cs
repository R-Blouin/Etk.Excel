using Etk.BindingTemplates.Definitions.EventCallBacks;
using System.ComponentModel.Composition;
using System.Reflection;
using Etk.Excel.Application;
using Etk.Tools.Reflection;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Definitions
{
    [Export(typeof(EventExcelCallbacksManager))]
    [Export(typeof(EventCallbacksManager))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class EventExcelCallbacksManager : EventCallbacksManager
    {
        protected override object InvokeNotDotNet(EventCallback callback, object[] parameters)
        {
            return ETKExcel.ExcelApplication.ExecuteVbaMAcro(callback.Ident, parameters);
        }
        public void RegisterSpecificCallBack()
        {
            MethodInfo methodInfo = TypeHelpers.GetMethod(typeof(ExcelApplication), "StaticShowHideColumns");
            SpecificEventCallback callback = new SpecificEventCallback("ETK_ShowHideColumns", "Manage show/hide columns on left double-click", methodInfo);
            callbackByIdent[callback.Ident] = callback;
        }

        public static void ShowHideColumns(ExcelInterop.Range targetedRange, int numberOfColumns)
        {
            ExcelApplication.StaticShowHideColumns(targetedRange, numberOfColumns);
        }
    }
}
