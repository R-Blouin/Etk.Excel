namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System.Xml.Serialization;

    /// <summary> Xml model type simple property definition</summary>
    public class XmlModelProperty
    {
        /// <summary> [Mandatory] Property Name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary> If not null, it will be used to override the real property 'Name'.</summary>
        [XmlAttribute]
        public string NameToUse        
        { get; set; }

        /// <summary> Property Description.</summary>
        [XmlAttribute]
        public string Description
        { get; set; }
    }
}
