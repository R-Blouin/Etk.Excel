namespace Etk.BindingTemplates.Definitions.EventCallBacks.XmlDefinitions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml.Serialization;

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
