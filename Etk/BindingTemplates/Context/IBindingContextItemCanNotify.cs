using System;
using System.ComponentModel;

namespace Etk.BindingTemplates.Context
{
    public interface IBindingContextItemCanNotify : IBindingContextItem 
    {
        Action<IBindingContextItem, object> OnPropertyChangedAction { get; set; }
        object OnPropertyChangedActionArgs{ get; set; }
        void OnPropertyChanged(object source, PropertyChangedEventArgs args);
    }
}
