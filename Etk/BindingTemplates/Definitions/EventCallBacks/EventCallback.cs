namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    using System;
    using System.Linq;
    using System.Reflection;
    using Etk.BindingTemplates.Context;
    using Etk.Excel.UI.Reflection;

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
        private EventCallback(string ident, string description, Type firstParameterType, MethodInfo toInvoke)
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
                EventCallback ret = EventCallbacksManager.GetCallback(methodName);
                if (ret == null)
                {
                    MethodInfo toInvoke = TypeHelpers.GetMethod(type, methodName);
                    if (toInvoke != null)
                        ret = new EventCallback(ident, description, type, toInvoke);
                }
                return ret;
            }
            catch (Exception ex)
            {
                throw new ArgumentException(string.Format("Method '{0}' not resolved:{1}", methodName, ex.Message), ex);
            }
        }
        #endregion

        #region public methods
        public bool Invoke(object sender, IBindingContextElement selectedContextElement, IBindingContextElement catchingContextElement)
        {
            object invokeTarget = Callback.IsStatic ? null : catchingContextElement.DataSource;
            int nbrParameters = Callback.GetParameters().Count();
            object[] parameters;

            switch (nbrParameters)
            {
                case 3:
                    parameters = new object[] { sender, selectedContextElement.DataSource, catchingContextElement.DataSource};
                break;
                case 2:
                    parameters = new object[] { selectedContextElement.DataSource, catchingContextElement.DataSource };
                break;
                default:
                    parameters = new object[] { selectedContextElement.DataSource };
                break;
            }

            Callback.Invoke(invokeTarget, parameters);
            return true; 
        }
        #endregion
    }
}
