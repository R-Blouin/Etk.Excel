using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Definitions.Decorators
{
    /// <summary> Manage the <see cref="Decorator"/> used in the current application</summary>
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class DecoratorsManager : IDisposable
    {
        private readonly ILogger log = Logger.Instance;
        private readonly Dictionary<string, Decorator> decoratorByIdent = new Dictionary<string, Decorator>();

        /// <summary> Return a <see cref="Decorator"/> given an ident</summary>
        /// <param name="ident">the ident of the <see cref="Decorator"/> to return</param>
        /// <returns>If found, the <see cref="Decorator"/>, if not: null</returns>
        public Decorator GetDecorator(string ident)
        {
            Decorator ret = null;
            if (!string.IsNullOrEmpty(ident))
            {
                lock ((decoratorByIdent as ICollection).SyncRoot)
                {
                    if (!decoratorByIdent.TryGetValue(ident, out ret))
                        throw new Exception($"Cannot find Decorator '{ident}'");
                }
            }
            return ret;
        }

        /// <summary> Register a <see cref="Decorator"/></summary>
        /// <param name="decorator">The <see cref="Decorator"/> to register</param>
        public void RegisterDecorator(Decorator decorator)
        {
            lock ((decoratorByIdent as ICollection).SyncRoot)
            {
                if(decorator != null)
                {
                    if (decoratorByIdent.ContainsKey(decorator.Ident))
                        log.LogFormat(LogType.Warn, "Decorator {0} already registred.", decorator.Ident ?? string.Empty);
                    decoratorByIdent[decorator.Ident] = decorator;
                }
            }
        }

        public void Dispose()
        {
            decoratorByIdent.Clear();
        }
    }
}
