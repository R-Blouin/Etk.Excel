namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System.Xml.Serialization;

    /// <summary> DefinitionToFilter of a model Type property that references another model type</summary>
    public class XmlModelLinkProperty
    {
        /// <summary> [Mandatory] Property Name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary> If not null, override the real property 'Name'.</summary>
        [XmlAttribute]
        public string NameToUse        
        { get; set; }

        /// <summary> Property Description.</summary>
        [XmlAttribute]
        public string Description
        { get; set; }

        /// <summary> For link peoperties: set the ModelAccessor to retrieve data.</summary>
        [XmlAttribute]
        public string Accessor
        { get; set; }

        /// <summary> For link peoperties: set the keys to invoke the 'ModelAccessor' or the 'ModelType' to retrieve the linked data.</summary>
        [XmlAttribute]
        public string Keys
        { get; set; }
    }
}
