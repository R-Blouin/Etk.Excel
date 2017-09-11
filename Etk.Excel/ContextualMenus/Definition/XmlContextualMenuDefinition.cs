using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Etk.Tools.Extensions;

namespace Etk.Excel.ContextualMenus.Definition
{
    [XmlRoot("ContextualMenu")]
    public class XmlContextualMenuDefinition : XmlContextualMenuPart
    {
        [XmlAttribute]
        public string Name
        { get; set; }

        [XmlAttribute]
        public string Caption
        { get; set; }

        [XmlAttribute]
        public string InsertBefore
        { get; set; }

        [XmlAttribute]
        public bool BeginGroup
        { get; set; }

        [XmlElement(ElementName = "MenuItem", Type = typeof(XmlContextualMenuItemDefinition))]
        [XmlElement(ElementName = "ContextualMenu", Type = typeof(XmlContextualMenuDefinition))]
        public List<XmlContextualMenuPart> Items
        { get; set; }

        public static XmlContextualMenuDefinition CreateInstance(string definitions)
        {
            XmlContextualMenuDefinition ret = null;
            definitions = definitions.EmptyIfNull().Trim();
            if (!string.IsNullOrEmpty(definitions))
            {
                try
                {
                    ret = definitions.Deserialize<XmlContextualMenuDefinition>();
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Cannot retrieve the contextual menu. {ex.Message}");
                }
            }
            return ret;
        }
    }
}
