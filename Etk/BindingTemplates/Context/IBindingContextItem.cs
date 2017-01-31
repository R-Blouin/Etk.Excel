using System;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Context
{
    public interface IBindingContextItem : IDisposable
    {
        //long Id { get; }
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
