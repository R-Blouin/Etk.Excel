namespace Etk.ModelManagement.Definitions.XmlDefinition
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;

    [XmlRoot("ModelConfiguration")]
    public class XmlModelConfiguration
    {
        /// <summary> Model Name</summary>
        [XmlAttribute]
        public string Name
        { get; set; }

        /// <summary> [Mandatory] Model accessorsNames.</summary>
        [XmlArray("ModelAccessorGroups")]
        [XmlArrayItem("ModelAccessorGroup", Type = typeof(XmlModelAccessorGroup))]
        public List<XmlModelAccessorGroup> ModelAccessorGroupDefinitions
        { get; set; }

        /// <summary> Model types.</summary>
        [XmlArray("ModelTypes")]
        [XmlArrayItem("ModelType", Type = typeof(XmlModelType))]
        public List<XmlModelType> TypeDefinitions
        { get; set; }

        #region static public methods
        static public XmlModelConfiguration CreateInstanceFromFile(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                    return null;

                if (! File.Exists(path))
                    throw new EtkException(string.Format("Cannot find the file '{0}", path));

                XmlModelConfiguration conf;
                using (FileStream reader = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(XmlModelConfiguration));
                    conf = serializer.Deserialize(reader) as XmlModelConfiguration;
                }
                return conf;
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("'XmlModelConfiguration' initialization failed:{0}", ex.InnerException == null ? ex.Message : ex.InnerException.Message));
            }
        }

        static public XmlModelConfiguration CreateInstanceFromXml(string xml)
        {
            try
            {
                if (string.IsNullOrEmpty(xml))
                    return null;

                using (StringReader reader = new StringReader(xml))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(XmlModelConfiguration));
                    XmlModelConfiguration conf = serializer.Deserialize(reader) as XmlModelConfiguration;
                    return conf;
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("'XmlModelConfiguration' initialization failed:{0}", ex.InnerException == null ? ex.Message : ex.InnerException.Message));
            }
        }
        #endregion
    }
}
