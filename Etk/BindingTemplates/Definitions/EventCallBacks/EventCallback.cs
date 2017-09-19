using System.Reflection;

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
    }
}
