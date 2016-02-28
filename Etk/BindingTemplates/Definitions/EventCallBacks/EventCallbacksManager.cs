namespace Etk.BindingTemplates.Definitions.EventCallBacks
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel.Composition;
    using Etk.BindingTemplates.Definitions.EventCallBacks.XmlDefinitions;
    using Etk.Excel.UI.Log;
    
    /// <summary> Manage the <see cref="Decorator"/> used in the current application</summary>
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class EventCallbacksManager : IDisposable
    {
        private ILogger log = Logger.Instance;
        private Dictionary<string, EventCallback> callbackByIdent = new Dictionary<string, EventCallback>();

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
                    {
                        EventCallback callback = EventCallback.CreateInstance(xmlcallback.Ident, xmlcallback.Description, null, xmlcallback.Method);
                        RegisterCallback(callback);
                    }
                }
            }
            catch (Exception ex)
            {
                string message = xml.Length > 350 ? xml.Substring(0, 350) + "..." : xml;
                throw new EtkException(string.Format("Cannot create Event Callbacks from xml '{0}':{1}", message, ex.Message), ex);
            }
        }


        /// <summary> Return a <see cref="EventCallback"/> given an ident</summary>
        /// <param name="ident">the ident of the <see cref="EventCallback"/> to return</param>
        /// <returns>If found, the <see cref="Decorator"/>, if not: null</returns>
        public EventCallback GetCallback(string ident)
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
        /// <param name="decorator">The <see cref="EventCallback"/> to register</param>
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

        public void Dispose()
        {
            callbackByIdent.Clear();
        }
    }
}
