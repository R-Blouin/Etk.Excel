namespace Etk.Excel.ContextualMenus.Definition
{
    using System.Collections.Generic;
    using System.Xml.Serialization;

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
    }
}
