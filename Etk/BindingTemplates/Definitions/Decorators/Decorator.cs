using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.EventCallBacks;

namespace Etk.BindingTemplates.Definitions.Decorators
{
    /// <summary> Contains a decorator definition.
    /// Decorators are used to change the default style of the <see cref="IBindingContextItem"/></summary>
    public abstract class Decorator
    {
        #region propertis and attributes
        /// <summary> Ident to use to reference the decorator</summary>
        public string Ident
        { get; private set; }

        /// <summary> Description</summary>
        public string Description
        { get; private set; }

        /// <summary> Call back to resolve the decorator</summary>
        public EventCallback Callback
        { get; private set; }

        //private IDecoratorProperty[] properties;
        ///// <summary> Properties (styles) that can be used by the decorator</summary>
        //public IEnumerable<IDecoratorProperty> Properties
        //{ get { return properties; } }
        #endregion

        #region .ctors
        public Decorator(string ident, string description, EventCallback callback)
        {
            Ident = ident;
            Description = description;
            Callback = callback;
        }
        #endregion

        #region public methods
        public abstract bool Resolve(object sender, IBindingContextItem contextItem);
        #endregion
    }
}
