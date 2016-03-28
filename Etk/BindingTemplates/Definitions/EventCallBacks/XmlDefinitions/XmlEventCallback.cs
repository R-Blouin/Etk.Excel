using System.Xml.Serialization;

namespace Etk.BindingTemplates.Definitions.EventCallBacks.XmlDefinitions
{
    public class XmlEventCallback
    {
        [XmlAttribute]
        public string Ident
        { get; set; }

        [XmlAttribute]
        public string Description
        { get; set; }

        [XmlAttribute]
        public string Method
        { get; set; }
    }
}
