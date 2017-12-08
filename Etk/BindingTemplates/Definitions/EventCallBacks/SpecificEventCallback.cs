using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;

namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    public class SpecificEventCallbackParameter
    {
        public bool IsSender { get; set; }
        public bool IsCatchingContextElement { get; set; }
        public bool IsCurrentContextItem { get; set; }

        public object ParameterValue { get; set; }
    }

    public class SpecificEventCallback : EventCallback
    {
        #region propertis and attributes
        public IEnumerable<SpecificEventCallbackParameter> Parameters
        { get; set; }
        #endregion

        #region .ctors and factory
        public SpecificEventCallback(string ident, string description, MethodInfo toInvoke) : base (ident, description, toInvoke)
        {}

        public SpecificEventCallback(SpecificEventCallback callback) : base(callback.Ident, callback.Description, callback.Callback)
        { }

        #endregion

        public void Invoke()
        {
            object[] methodParams = null;
            if (Parameters != null && Parameters.Any())
            {
                methodParams = new object[Parameters.Count()];
                int i = 0;
                foreach (SpecificEventCallbackParameter param in Parameters)
                    methodParams[i++] = param.ParameterValue;
            }
            Callback.Invoke(null, methodParams);
        }

        public override void Invoke(object sender, IBindingContextElement catchingContextElement, IBindingContextItem currentContextItem)
        {
            object[] methodParams = null;
            if (Parameters != null && Parameters.Any())
            {
                methodParams = new object[Parameters.Count()];
                int i = 0;
                foreach (SpecificEventCallbackParameter param in Parameters)
                {
                    if (param.IsSender)
                    {
                        methodParams[i++] = sender;
                        continue;
                    }
                    if (param.IsCatchingContextElement)
                    {
                        methodParams[i++] = catchingContextElement;
                        continue;
                    }
                    if (param.IsCurrentContextItem)
                    {
                        methodParams[i++] = currentContextItem;
                        continue;
                    }
                    methodParams[i++] = param.ParameterValue;
                }
            }
            Callback.Invoke(null, methodParams);
        }
    }
}
