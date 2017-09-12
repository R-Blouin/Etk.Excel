using System;
using System.Linq;
using System.Reflection;
using Etk.Tools.Reflection;

namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    public class EventCallback
    {
        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager
        {
            get
            {
                if (eventCallbacksManager == null)
                    eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>();
                return eventCallbacksManager;
            }
        }


        #region propertis and attributes
        /// <summary> Ident to use to reference the decorator</summary>
        public string Ident
        { get; private set; }

        /// <summary> Description</summary>
        public string Description
        { get; private set; }

        /// <summary> Method info to invoke</summary>
        public MethodInfo Callback
        { get; private set; }

        #endregion

        #region .ctors and factory
        private EventCallback(string ident, string description, MethodInfo toInvoke)
        {
            ParameterInfo[] parameters = toInvoke.GetParameters();
            if (toInvoke.ReturnType != typeof(void) || parameters == null || parameters.Count() > 3)
                throw new EtkException("Method dataAccessor must be 'void MethodName([Range <selected range>,] [object <catching object>,] object <selected object>)'");

            Ident = ident;
            Description = description;
            Callback = toInvoke;
        }

        public static EventCallback CreateInstance(string ident, string description, Type type, string methodName)
        {
            try
            {
                EventCallback ret = EventCallbacksManager.GetCallback(ident);
                if (ret == null)
                {
                    MethodInfo toInvoke = TypeHelpers.GetMethod(type, methodName);
                    if (toInvoke != null)
                        ret = new EventCallback(ident, description, toInvoke);
                }
                return ret;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Method '{methodName}' not resolved:{ex.Message}");
            }
        }
        #endregion
    }
}
