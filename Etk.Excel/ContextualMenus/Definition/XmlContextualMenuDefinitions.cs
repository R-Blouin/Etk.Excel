namespace Etk.Excel.ContextualMenus.Definition
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml.Serialization;
    using Etk.Excel.UI.Extensions;

    [XmlRoot("ContextualMenus")]
    public class XmlContextualMenuDefinitions
    {
        [XmlElement(ElementName = "ContextualMenu", Type = typeof(XmlContextualMenuDefinition))]
        public List<XmlContextualMenuDefinition> ContextualMenus
        { get; set; }

        static public XmlContextualMenuDefinitions CreateInstance(string definitions)
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
                    string message = string.Format("Cannot retrieve the contextual menus. {0}", ex.Message);
                    throw new EtkException(message, ex);
                }
            }
            return ret;
        }
    }
}
