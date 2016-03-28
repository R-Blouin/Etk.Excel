using System;
using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Context
{
    public class BindingContextItemCanNotify : BindingContextItem, IBindingContextItemCanNotify 
    {
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;
        
        public Action<IBindingContextItem, object> OnPropertyChangedAction
        { get; set; }

        public object OnPropertyChangedActionArgs
        { get; set; }

        public BindingContextItemCanNotify(IBindingContextElement parent, IBindingDefinition bindingDefinition)
                                          : base(parent, bindingDefinition)
        {
            CanNotify = true;

            objectsToNotify = bindingDefinition.GetObjectsToNotify(DataSource);
            if (objectsToNotify != null)
            {
                foreach (INotifyPropertyChanged obj in objectsToNotify)
                    obj.PropertyChanged += OnPropertyChanged;
            }
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (objectsToNotify != null && OnPropertyChangedAction != null)
            {
                if (BindingDefinition.MustNotify(DataSource, source, args))
                    OnPropertyChangedAction(this, OnPropertyChangedActionArgs);
            }
        }

        public override void RealDispose()
        {
            OnPropertyChangedAction = null;

            if (objectsToNotify != null)
            {
                foreach (INotifyPropertyChanged obj in objectsToNotify)
                    obj.PropertyChanged -= OnPropertyChanged;
                objectsToNotify = null;
            }
        }
    }
}
