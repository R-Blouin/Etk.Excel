namespace Etk.BindingTemplates.Definitions.EventCallBacks.XmlDefinitions
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;
    using Etk.Excel.UI.Extensions;

    [XmlRoot("EventCallbacks")]
    public class XmlEventCallbacks
    {
        [XmlElement(ElementName = "EventCallback", Type = typeof(XmlEventCallback))]
        public List<XmlEventCallback> Callbacks
        { get; set; }

        public static XmlEventCallbacks CreateInstance(string xml)
        {
            XmlEventCallbacks ret = null;
            xml = xml.EmptyIfNull().Trim();
            if (!string.IsNullOrEmpty(xml))
            {
                try
                {
                    ret = xml.Deserialize<XmlEventCallbacks>();
                }
                catch (Exception ex)
                {
                    string def = xml.EmptyIfNull().Trim();
                    if (def.Length > 150)
                        def = def.Substring(0, 149) + "...";

                    string message = string.Format("Cannot retrieve the Event Callback from '{0}'. {1}", def, ex.Message);
                    throw new EtkException(message, ex);
                }
            }
            return ret;
        }
    }
}