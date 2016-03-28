using System.Collections.Generic;
using System.Xml.Serialization;

namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    /// <summary> Model Accessor groups.</summary>
    public class XmlModelAccessorGroup
    {

        /// <summary> [Mandatory] Accessor Name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary>Accessor Description</summary>
        [XmlAttribute]
        public string Description
        { get; set; }

        /// <summary> Model Accessors.</summary>
        [XmlElement("ModelAccessor", Type = typeof(XmlModelAccessor))]
        public List<XmlModelAccessor> Accessors
        { get; set; }
    }
}
