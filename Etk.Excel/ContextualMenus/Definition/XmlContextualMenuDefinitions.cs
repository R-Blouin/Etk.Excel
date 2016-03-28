using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Etk.Tools.Extensions;

namespace Etk.Excel.ContextualMenus.Definition
{
    [XmlRoot("ContextualMenus")]
    public class XmlContextualMenuDefinitions
    {
        [XmlElement(ElementName = "ContextualMenu", Type = typeof(XmlContextualMenuDefinition))]
        public List<XmlContextualMenuDefinition> ContextualMenus
        { get; set; }

        public static XmlContextualMenuDefinitions CreateInstance(string definitions)
        {
            XmlContextualMenuDefinitions ret = null;
            definitions = definitions.EmptyIfNull().Trim();
            if (!string.IsNullOrEmpty(definitions))
            {
                try
                {
                    ret = definitions.Deserialize<XmlContextualMenuDefinitions>();
                }
                catch (Exception ex)
                {
                    throw new EtkException(string.Format("Cannot retrieve the contextual menus. {0}", ex.Message));
                }
            }
            return ret;
        }
    }
}
