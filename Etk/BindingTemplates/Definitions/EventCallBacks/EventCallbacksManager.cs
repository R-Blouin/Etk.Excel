using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.EventCallBacks.XmlDefinitions;
using Etk.BindingTemplates.Views;
using Etk.Tools.Log;
using Etk.Tools.Reflection;
using System.Linq;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    /// <summary> Manage the <see cref="EventCallback"/> used in the current application</summary>
    public class EventCallbacksManager : IDisposable
    {
        protected readonly ILogger log = Logger.Instance;
        protected readonly Dictionary<string, EventCallback> callbackByIdent = new Dictionary<string, EventCallback>();

        #region public methods
        /// <summary>Register event callbacks from xml definitions</summary>
        /// <param name="xml">The xml that contains the callback definitions </param>
        public void RegisterEventCallbacksFromXml(string xml)
        {
            try
            {
                XmlEventCallbacks xmlCallbacks = XmlEventCallbacks.CreateInstance(xml);
                if (xmlCallbacks == null)
                    return;

                if (xmlCallbacks.Callbacks != null)
                {
                    foreach (XmlEventCallback xmlcallback in xmlCallbacks.Callbacks)
                        RegisteredCallback(xmlcallback.Ident, xmlcallback.Description, null, xmlcallback.Method);
                }
            }
            catch (Exception ex)
            {
                string message = xml.Length > 350 ? xml.Substring(0, 350) + "..." : xml;
                throw new EtkException($"Cannot create Event Callbacks from xml '{message}':{ex.Message}");
            }
        }

        /// <summary> Return a <see cref="EventCallback"/> given an ident of a callback previously registered</summary>
        /// <param name="ident">the ident of the <see cref="EventCallback"/> to return</param>
        /// <returns>If found, the <see cref="EventCallback"/>, if not: null</returns>
        public EventCallback GetRegisteredCallback(string ident)
        {
            EventCallback ret = null;
            if (!String.IsNullOrEmpty(ident))
            {
                lock ((callbackByIdent as ICollection).SyncRoot)
                {
                    callbackByIdent.TryGetValue(ident, out ret);
                }
            }
            return ret;
        }

        /// <summary> Register a <see cref="EventCallback"/></summary>
        /// <param name="callback">The <see cref="EventCallback"/> to register</param>
        public void RegisterCallback(EventCallback callback)
        {
            lock ((callbackByIdent as ICollection).SyncRoot)
            {
                if (callback != null)
                {
                    if (callbackByIdent.ContainsKey(callback.Ident))
                        log.LogFormat(LogType.Warn, "EventCallback {0} already registred.", callback.Ident ?? string.Empty);
                    callbackByIdent[callback.Ident] = callback;
                }
            }
        }

        public EventCallback RetrieveCallback(ITemplateDefinition templateDefinition, string callbackName)
        {
            EventCallback ret = null;

            if (callbackName.StartsWith("$")) // The callback is not a .Net one
                ret = new EventCallback(callbackName.TrimStart('$'), null, null);
            else if (callbackName.StartsWith("*"))  // The callback was registred (from xml)
                ret = GetRegisteredCallback(callbackName);
            else
            {
                MethodInfo methodInfo;

                string methodName = callbackName;
                string[] parts = callbackName.Split(',');
                if (parts.Count() == 1) // The callback is a member of the 'templateDefinition.MainBindingDefinition.BindingType' class
                    ret = GetCallBackFromMainBindingDefinition(templateDefinition, parts[0]);
                if (parts.Count() == 3) // assembly, type and nam are supplied
                {
                    if(string.IsNullOrEmpty(parts[0]) && string.IsNullOrEmpty(parts[1]))
                        ret = GetCallBackFromMainBindingDefinition(templateDefinition, parts[2]);
                    else
                    {
                        methodInfo = TypeHelpers.GetMethod(null, callbackName);
                        ret = new EventCallback(null, null, methodInfo);
                    }
                }
            }

            if (ret == null)
                throw new Exception($"Cannot find the callback '{callbackName}'");

            return ret;
        }

        public void Invoke(EventCallback callback, object sender, IBindingContextElement catchingContextElement, IBindingContextItem currentContextItem)
        {
            if (callback.IsNotDotNet)
                InvokeNotDotNet(callback, sender, catchingContextElement, currentContextItem);
            else
            {
                MethodInfo methodInfo = callback.Callback;
                object invokeTarget = methodInfo.IsStatic ? null : catchingContextElement.DataSource;
                int nbrParameters = methodInfo.GetParameters().Length;

                if (nbrParameters > 3)
                    throw new Exception($"Method info '{methodInfo.Name}' signature is not correct");

                object[] parameters;
                switch (nbrParameters)
                {
                    //case 4:
                    //    parameters = new object[] { catchingContextElement, catchingContextElement.DataSource, currentContextItem, currentContextItem.DataSource };
                    //    break;
                    case 3:
                        parameters = new object[] { sender, catchingContextElement.DataSource, currentContextItem.ParentElement.DataSource };
                        break;
                    case 2:
                        if (methodInfo.GetParameters()[0].ParameterType == typeof(ITemplateView))
                            parameters = new object[] { catchingContextElement.ParentPart.ParentContext.Owner, catchingContextElement.DataSource };
                        else
                            parameters = new object[] { catchingContextElement.DataSource, currentContextItem.ParentElement.DataSource };
                        break;
                    case 1:
                        parameters = new object[] { catchingContextElement.DataSource };
                        break;
                    default:
                        parameters = null;
                        break;
                }
                methodInfo.Invoke(invokeTarget, parameters);
            }
        }

        public void Dispose()
        {
            callbackByIdent.Clear();
        }
        #endregion

        protected virtual void InvokeNotDotNet(EventCallback callback, object sender, IBindingContextElement catchingContextElement, IBindingContextItem currentContextItem)
        {}

        /// <summary></summary>
        private  void RegisteredCallback(string ident, string description, Type type, string methodName)
        {
            try
            {
                EventCallback callback;
                if(callbackByIdent.TryGetValue(ident, out callback))
                {
                    MethodInfo toInvoke = null;
                    if(methodName != null && ! methodName.StartsWith("$"))
                        toInvoke =  TypeHelpers.GetMethod(type, methodName);
                    callback = new EventCallback(ident, description, toInvoke);
                    RegisterCallback(callback);
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Method '{methodName??string.Empty}' not resolved:{ex.Message}");
            }
        }

        private EventCallback GetCallBackFromMainBindingDefinition(ITemplateDefinition templateDefinition, string methodName)
        {
            MethodInfo methodInfo;
            EventCallback ret = null;
            Type inType = null;
            if (templateDefinition != null && templateDefinition.MainBindingDefinition != null && templateDefinition.MainBindingDefinition.BindingType != null)
            {
                inType = templateDefinition.MainBindingDefinition.BindingType;
                methodInfo = TypeHelpers.GetMethod(inType, methodName);
                ret = new EventCallback(null, null, methodInfo);
            }
            return ret;
        }
    }
}
