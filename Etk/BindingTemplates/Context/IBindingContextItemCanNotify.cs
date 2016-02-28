namespace Etk.BindingTemplates.Context
{
    using System;
    using System.ComponentModel;

    public interface IBindingContextItemCanNotify : IBindingContextItem 
    {
        Action<IBindingContextItem, object> OnPropertyChangedAction { get; set; }
        object OnPropertyChangedActionArgs{ get; set; }
        void OnPropertyChanged(object source, PropertyChangedEventArgs args);
    }
}
