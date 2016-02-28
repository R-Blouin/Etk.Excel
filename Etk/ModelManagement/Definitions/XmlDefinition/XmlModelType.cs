namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System.Collections.Generic;
    using System.Xml.Serialization;

    /// <summary> Xml model type definition </summary>
    public class XmlModelType
    {
        /// <summary> Model Type name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary> Model Type Description.</summary>
        [XmlAttribute]
        public string Description
        { get; set; }

        /// <summary> Model Type underlying .Net type.</summary>
        [XmlAttribute]
        public string Type
        { get; set; }

        ///// <summary> Model Type reference .</summary>
        //[XmlAttribute]
        //public string Reference
        //{ get; set; }

        /// <summary> Model Type default properties.</summary>
        [XmlElement("Property", Type = typeof(XmlModelProperty))]
        public List<XmlModelProperty> Properties
        { get; set; }

        /// <summary> Model Type properties from underlying .Net type to ignore (to exclude).</summary>
        [XmlElement("PropertyToIgnore")]
        public List<string> PropertiesToIgnore
        { get; set; }

        /// <summary> Model Type property that reference another model type.</summary>
        [XmlElement("LinkProperty", Type = typeof(XmlModelLinkProperty))]
        public List<XmlModelLinkProperty> LinkProperties
        { get; set; }


        /// <summary> Default views supplied by the model.</summary>
        [XmlElement("View", Type = typeof(XmlModelView))]
        public List<XmlModelView> Views
        { get; set; }
    }
}
