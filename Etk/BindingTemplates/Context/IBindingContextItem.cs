namespace Etk.BindingTemplates.Context
{
    using System;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Views;

    public interface IBindingContextItem : IDisposable
    {
        long Id { get; }
        string Name { get; }
        IBindingContextElement ParentElement { get; }
        IBindingDefinition BindingDefinition { get; }
        object DataSource { get; }
        bool CanNotify { get; }

        bool IsDisposed { get; }
        object ResolveBinding();
        bool UpdateDataSource(object data, out object retValue);
    }
}
