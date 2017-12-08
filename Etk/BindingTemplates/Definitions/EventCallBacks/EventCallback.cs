using System;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Views;

namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    public class EventCallback
    {
        #region propertis and attributes
        /// <summary> Ident to use to reference the decorator</summary>
        public string Ident
        { get; private set; }

        /// <summary> If true, the callback is not a Not Net one</summary>
        public bool IsNotDotNet
        { get; private set; }


        /// <summary> Description</summary>
        public string Description
        { get; private set; }

        /// <summary> Method info to invoke</summary>
        public MethodInfo Callback
        { get; private set; }

        #endregion

        #region .ctors and factory
        public EventCallback(string ident, string description, MethodInfo toInvoke)
        {
            //if (toInvoke != null)
            //{
            //    ParameterInfo[] parameters = toInvoke.GetParameters();
            //    if (toInvoke.ReturnType != typeof(void) || parameters == null || parameters.Count() > 3)
            //        throw new EtkException("Method dataAccessor must be 'void MethodName([Range <selected range>,] [object <catching object>,] object <selected object>)'");
            //}

            Ident = ident;
            Description = description;
            Callback = toInvoke;
            IsNotDotNet = toInvoke == null;
        }
        #endregion

        public virtual void Invoke(object sender, IBindingContextElement catchingContextElement, IBindingContextItem currentContextItem)
        {
            object invokeTarget = Callback.IsStatic ? null : catchingContextElement.DataSource;
            int nbrParameters = Callback.GetParameters().Length;

            if (nbrParameters > 3)
                throw new Exception($"Method info '{Callback.Name}' signature is not correct");

            object[] parameters;
            switch (nbrParameters)
            {
                //case 4:
                //    parameters = new object[] { catchingContextElement, catchingContextElement.DataSource, currentContextItem, currentContextItem.DataSource };
                //    break;
                case 3:
                    parameters = new[] { sender, catchingContextElement.DataSource, currentContextItem.ParentElement.DataSource };
                    break;
                case 2:
                    if (Callback.GetParameters()[0].ParameterType == typeof(ITemplateView))
                        parameters = new[] { catchingContextElement.ParentPart.ParentContext.Owner, catchingContextElement.DataSource };
                    else
                        parameters = new[] { catchingContextElement.DataSource, currentContextItem.ParentElement.DataSource };
                    break;
                case 1:
                    parameters = new[] { catchingContextElement.DataSource };
                    break;
                default:
                    parameters = null;
                    break;
            }
            Callback.Invoke(invokeTarget, parameters);
        }
    }
}
