namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml.Serialization;

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
