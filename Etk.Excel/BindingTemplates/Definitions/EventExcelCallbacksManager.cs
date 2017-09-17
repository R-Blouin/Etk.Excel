using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;

namespace Etk.Excel.BindingTemplates.Definitions
{
    [Export(typeof(EventExcelCallbacksManager))]
    [Export(typeof(EventCallbacksManager))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class EventExcelCallbacksManager : EventCallbacksManager
    {
        protected override void InvokeNotDotNet(EventCallback callback, object sender, IBindingContextElement catchingContextElement, IBindingContextItem currentContextItem)
        {
            ExcelTemplateDefinition templateDefinition = catchingContextElement.ParentPart.ParentContext.TemplateDefinition as ExcelTemplateDefinition;
            try
            {
                ETKExcel.ExcelApplication.ExecuteVbaMAcro(callback.Ident, new[] { catchingContextElement?.DataSource, currentContextItem.BindingDefinition.Name });
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode != (int) SpecificException.DISP_E_UNKNOWNNAME)
                    throw;
            }
        }
    }
}
