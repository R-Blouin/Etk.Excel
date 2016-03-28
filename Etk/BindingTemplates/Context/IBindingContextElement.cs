using System;
using System.Collections.Generic;

namespace Etk.BindingTemplates.Context
{
    public interface IBindingContextElement: IDisposable
    {
        IBindingContextPart ParentPart { get; }
        object DataSource { get; }
        object Element { get; }
        int Index { get; }

        List<IBindingContextItem> BindingContextItems { get; }
        List<IBindingContext> LinkedBindingContexts { get; }
    }
}
