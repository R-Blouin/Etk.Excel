using Etk.BindingTemplates.Definitions.EventCallBacks;
using System.ComponentModel.Composition;

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
    }
}
