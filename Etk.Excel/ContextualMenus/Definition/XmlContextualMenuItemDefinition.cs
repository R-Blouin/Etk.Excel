using System.Xml.Serialization;

namespace Etk.Excel.ContextualMenus.Definition
{
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
