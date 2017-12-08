using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Context;
using System.Runtime.InteropServices;
using System.Reflection;
using Etk.Tools.Log;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionOptional : BindingDefinition
    {
        private bool comNameIsInvalid;
        private readonly ILogger log = Logger.Instance;
        private static readonly object syncObj = new object();

        private readonly Dictionary<Type, IBindingDefinition> bindingDefinitionByType = new Dictionary<Type, IBindingDefinition>();
        
        public override string Name => string.IsNullOrEmpty(DefinitionDescription.Name) ? DefinitionDescription.BindingExpression : DefinitionDescription.Name;

        #region .ctors and factories
        private BindingDefinitionOptional(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        { }
        
        public static BindingDefinitionOptional CreateInstance(BindingDefinitionDescription definitionDescription)
        {
            BindingDefinitionOptional definition = new BindingDefinitionOptional(definitionDescription) { IsOptional = true };
            return definition;
        }
        #endregion

        #region public methods
        public override object UpdateDataSource(object dataSource, object data)
        {
            try
            {
                if (dataSource == null || comNameIsInvalid)
                    return null;

                if (! IsReadOnly)
                {
                    if (Marshal.IsComObject(dataSource) && !comNameIsInvalid)
                    {
                        //lock (syncObj) // we need to synchro Com exec
                        //{
                            dataSource.GetType().InvokeMember(Name, BindingFlags.Default | BindingFlags.SetProperty, null, dataSource, new[] { data }, null);
                        //}
                    }
                }
                return ResolveBinding(dataSource);
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == (int) SpecificException.DISP_E_UNKNOWNNAME)
                    comNameIsInvalid = true;
                return null;
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "'UpdateDataSource' failed for BindingExpression '{0}', value '{1}': {2}", BindingExpression, data?.ToString() ?? string.Empty, ex.Message);
                return ResolveBinding(dataSource);
            }
        }

        public override object ResolveBinding(object dataSource)
        {
            if (dataSource != null)
            {
                Type type = dataSource.GetType();
                if (Marshal.IsComObject(dataSource) && !comNameIsInvalid)
                {
                    try
                    {
                        //lock (syncObj) // we need to synchro Com exec
                        //{
                        object ret = type.InvokeMember(Name, BindingFlags.Default | BindingFlags.GetProperty, null, dataSource, null, null);
                        //}
                    }
                    catch (COMException ex)
                    {
                        if (ex.ErrorCode == (int)SpecificException.DISP_E_UNKNOWNNAME)
                            comNameIsInvalid = true;
                    }
                }
            }
            return null;
        }

        public IBindingDefinition CreateRealBindingDefinition(Type type)
        { 
            IBindingDefinition definition;
            if (!bindingDefinitionByType.TryGetValue(type, out definition))
            {
                definition = BindingDefinitionFactory.CreateInstance(type, DefinitionDescription) ?? this;
                bindingDefinitionByType[type] = definition;
            }
            return definition;
        }

        public override IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            BindingContextItem ret;
            if (parent.DataSource == null)
                ret = new BindingContextItem(parent, this);
            else
            {
                IBindingDefinition realBindingDefinition = CreateRealBindingDefinition(parent.DataSource.GetType());
                ret = realBindingDefinition.CanNotify ? new BindingContextItemCanNotify(parent, realBindingDefinition) 
                                                        : new BindingContextItem(parent, realBindingDefinition);
            }
            ret.Init();
            return ret;
        }
        #endregion
    }
}
