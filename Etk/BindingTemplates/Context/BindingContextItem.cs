namespace Etk.BindingTemplates.Context
{
    using System;
    using System.Threading;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Views;
    using Etk.Excel.UI.Extensions;
    using Etk.Excel.UI.Log;

    public class BindingContextItem : IBindingContextItem
    {
        #region attributes and properties
        private static long CurrentId = 0;
        
        private ILogger log = Logger.Instance;

        public IBindingContextElement ParentElement
        { get; private set; }

        public long Id
        { get; private set; }

        public string Name
        { get; private set; }

        public IBindingDefinition BindingDefinition
        { get; private set; }

        public object DataSource
        { get { return ParentElement.DataSource; } }

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
            if (BindingDefinition != null  && BindingDefinition.IsBoundWithData)
                Id = Interlocked.Increment(ref CurrentId);
            else
                Id = -1;

             if (ParentElement.Element != null && BindingDefinition != null && BindingDefinition.IsBoundWithData)
            {
                BindingTypeProperty property = ParentElement.ParentPart.ParentContext.TemplateDefinition.BindingType.PropertyByName[BindingDefinition.Name];
                property.SetMethod.Invoke(ParentElement.Element, new object[] { ResolveBinding() });
            }
        }
        #endregion

        #region public methods
        virtual public object ResolveBinding()
        {
            try
            {
                return BindingDefinition == null ? null : BindingDefinition.ResolveBinding(this.DataSource);
            }
            catch (Exception ex)
            {
                string message = string.Format("Can't resolve binding for 'BindingDefinition' '{0}': {1}", Name, ex.Message.EmptyIfNull());
                log.LogException(LogType.Error, ex, message);
                return "##Binding Error##";
            }
        }

        virtual public bool UpdateDataSource(object data, out object retValue)
        {
            if (this.BindingDefinition != null)
            {
                retValue = this.BindingDefinition.UpdateDataSource(this.DataSource, data);
                return true;
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
