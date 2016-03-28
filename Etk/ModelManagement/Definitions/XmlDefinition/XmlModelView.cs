using System.Xml.Serialization;

namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    /// <summary> Xml model view definition. Describe a view on a model type</summary>
    public class XmlModelView
    {
        /// <summary> Model View name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary> Model View Description.</summary>
        [XmlAttribute]
        public string Description
        { get; set; }

        /// <summary> If true, then this view is the default one for its model type.</summary>
        [XmlAttribute]
        public bool IsDefault
        { get; set; }

        /// <summary> Contain the name of the accessor the access the data. If none is supplied, then first accessor of the underlying model type will be used.</summary>
        [XmlAttribute]
        public string Accessor
        { get; set; }

        /// <summary>[Mandatory] Contain the view property list.</summary>
        [XmlText]
        public string Properties
        { get; set; }
    }
}
