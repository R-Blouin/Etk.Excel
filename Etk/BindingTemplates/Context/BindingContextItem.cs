using System;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Tools.Extensions;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Context
{
    public class BindingContextItem : IBindingContextItem
    {
        #region attributes and properties
        //private static long CurrentId = 0;
        
        private readonly ILogger log = Logger.Instance;

        public IBindingContextElement ParentElement
        { get;  }

        //public long Id
        //{ get; private set; }

        public string Name
        { get;  }

        public IBindingDefinition BindingDefinition
        { get; }

        public object DataSource => ParentElement.DataSource;

        public bool CanNotify
        { get; protected set; }

        public bool IsDisposed
        { get; private set; }
        #endregion

        #region .ctors
        public BindingContextItem(IBindingContextElement parent, IBindingDefinition bindingDefinition)
        {
            ParentElement = parent;
            BindingDefinition = bindingDefinition;
            CanNotify = false;
            Name = BindingDefinition == null ? string.Empty : BindingDefinition.Name;
        }
        #endregion

        #region internal methods
        internal void Init()
        {
            //if (BindingDefinition != null  && BindingDefinition.IsBoundWithData)
            //    Id = Interlocked.Increment(ref CurrentId);
            //else
            //    Id = -1;

            if (ParentElement.Element != null && BindingDefinition != null && BindingDefinition.IsBoundWithData)
            {
                BindingTypeProperty property = ParentElement.ParentPart.ParentContext.TemplateDefinition.BindingType.PropertyByName[BindingDefinition.Name];
                property.SetMethod.Invoke(ParentElement.Element, new [] { ResolveBinding() });
            }
        }
        #endregion

        #region public methods
        public virtual object ResolveBinding()
        {
            try
            {
                return BindingDefinition?.ResolveBinding(DataSource);
            }
            catch (Exception ex)
            {
                string message = $"Can't resolve binding for 'BindingDefinition' '{Name}': {ex.Message.EmptyIfNull()}";
                log.LogException(LogType.Error, ex, message);
                return "##Binding Error##";
            }
        }

        public virtual bool UpdateDataSource(object data, out object retValue)
        {
            if (BindingDefinition != null)
            {
                retValue = BindingDefinition.UpdateDataSource(DataSource, data);
                return BindingDefinition.IsReadOnly;
            }
            else
            { 
                retValue = null;
                return false;
            }
        }

        public void Dispose()
        {
            if (!IsDisposed)
            {
                IsDisposed = true;
                RealDispose();
            }
        }

        public virtual void RealDispose()
        {}
        #endregion
    }
}
