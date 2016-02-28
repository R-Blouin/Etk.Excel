namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System.Xml.Serialization;

    public class XmlModelAccessor
    {
        /// <summary> Accessor Ident. If not supplied, then the name will be used.</summary>
        [XmlAttribute]
        public string Ident
        { get; set; }

        /// <summary> [Mandatory] Accessor Name.</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary>Accessor Description</summary>
        [XmlAttribute]
        public string Description
        { get; set; }

        /// <summary> [Mandatory] Method to be invoked</summary>
        [XmlAttribute]
        public string Method
        { get; set; }

        /// <summary> Name of the Model type returned by the 'Method' invocation.
        /// If not supplied then the return type is deduced form the return type of the 'Method' property</summary>
        [XmlAttribute]
        public string ReturnModelType
        { get; set; }

        /// <summary> Only valid if 'DataAccessor' is settled. Can be 'Static, Singleton'</summary>
        [XmlAttribute]
        public string InstanceType
        { get; set; }

        /// <summary> [Mandatory if InstanceType = Singleton] Name (methdo or property) of the singleton accessor. </summary>
        [XmlAttribute]
        public string InstanceName
        { get; set; }
    }
}
