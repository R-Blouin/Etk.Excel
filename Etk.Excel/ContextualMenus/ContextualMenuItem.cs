using System;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Tools.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.ContextualMenus
{
    class ContextualMenuItem : IContextualMenuItem
    {
        #region attributes and properties
        public string Caption
        { get; private set; }

        public bool BeginGroup
        { get; private set; }

        public int FaceId
        { get; private set; }

        public MethodInfo MethodInfo
        { get; set; }

        public System.Action Action
        { get; private set; }
        #endregion

        #region .ctors
        public ContextualMenuItem(string caption, bool beginGroup, MethodInfo methodInfo, int faceId)
        {
            try
            {
                if (string.IsNullOrEmpty(caption))
                    throw new EtkException("'caption' cannot be null");

                Caption = caption;
                BeginGroup = beginGroup;
                MethodInfo = methodInfo;
                FaceId = faceId;
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Cannot create contextual menu '{0}': {1}", caption.EmptyIfNull(), ex.Message));
            }
        }
        #endregion

        public void SetAction(ExcelInterop.Range range, IBindingContextElement catchingContextElement, IBindingContextElement currentContextItem)
        {
            //ExcelTemplateDefinitionPart currentTemplateDefinition = catchingContextElement.ParentPart.TemplateDefinitionPart as ExcelTemplateDefinitionPart;
            if (MethodInfo != null)
            {
                object concernedObject = MethodInfo.IsStatic ? null : catchingContextElement.DataSource;
                int nbrParameters = MethodInfo.GetParameters().Count();
                if (nbrParameters == 3)
                    Action = () => MethodInfo.Invoke(concernedObject, new object[] { range, catchingContextElement.DataSource, currentContextItem.DataSource });
                else if (nbrParameters == 2)
                    Action = () => MethodInfo.Invoke(concernedObject, new object[] { catchingContextElement.DataSource, currentContextItem.DataSource });
                else if (nbrParameters == 1)
                    Action = () => MethodInfo.Invoke(concernedObject, new object[] { catchingContextElement.DataSource });
                else
                    Action = () => MethodInfo.Invoke(concernedObject, null);
            }
        }

        public void SetAction(ExcelInterop.Range range)
        {
            if (MethodInfo != null)
                Action = () => MethodInfo.Invoke(null, new object[] { range});
        }

        public void SetAction()
        {
            if (MethodInfo != null)
                Action = () => MethodInfo.Invoke(null, null);
        }

        public void SetAction(System.Action action)
        {
            if (action != null)
                Action = action;
        }
    }
}
