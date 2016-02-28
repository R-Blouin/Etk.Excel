namespace Etk.Excel.ContextualMenus.Definition
{
    using System.Xml.Serialization;

    public class XmlContextualMenuItemDefinition : XmlContextualMenuPart
    {
        [XmlAttribute]
        public string Caption
        { get; set; }

        [XmlAttribute]
        public bool BeginGroup
        { get; set; }

        [XmlAttribute]
        public int FaceId
        { get; set; }

        [XmlAttribute]
        public string Action
        { get; set; }
    }
}
