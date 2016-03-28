using System;
using System.Configuration;
using System.Xml;
using System.Xml.Serialization;

namespace Etk.Tools.Configuration
{
    /// <summary>
    /// Deserialise a section of the application config file into an instance of a given class.
    /// <para>Below, a configuration example of the App.Config</para>
    /// <example>
    /// <para>&lt;configuration&gt;</para>
    /// <para>...</para>
    /// <para>&lt;configSections&gt;</para>
    /// <para>&#160;&#160;&lt;section name="SectionName" type="Etk.Configuration.XmlSerializerSectionHandler, Etk"/&gt;</para>
    /// <para>&lt;/configSections&gt;</para>
    /// <para>...</para>  
    /// <para>&lt;SectionName type="FullTargetType, AssemblyContainingTheTargetType"&gt;</para>
    /// <para>&#160;&#160;XmlOfTheClassInstanceToDeserialize</para>
    /// <para>&lt;/SectionName&gt;</para> 
    /// <para>...</para>
    /// <para>&lt;/configuration&gt;</para>
    /// </example>
    /// </summary>
    public class XmlSerializerSectionHandler : IConfigurationSectionHandler
    {
        /// <summary>
        /// <see cref="M:System.Configuration.IConfigurationSectionHandler"/>
        /// </summary>
        /// <param name="parent"><see cref="M:System.Configuration.IConfigurationSectionHandler"/></param>
        /// <param name="configContext"><see cref="M:System.Configuration.IConfigurationSectionHandler"/></param>
        /// <param name="section"><see cref="M:System.Configuration.IConfigurationSectionHandler"/></param>
        /// <returns>un object having the 'Type' of the section.</returns>
        public object Create(object parent, object configContext, XmlNode section)
        {
            try
            {
                XmlAttribute typeAttribute = section.Attributes["UnderlyingType"];
                if (typeAttribute == null)
                    throw new ArgumentException(string.Format("the attribut 'UnderlyingType', section '{0}' is not set", section.Name));

                Type type = Type.GetType(typeAttribute.InnerText);
                XmlSerializer xmlSerializer = new XmlSerializer(type);
                return xmlSerializer.Deserialize(new XmlNodeReader(section));
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("XmlSerializerSectionHandler failed, section '{0}': {1}", section.Name, ex.Message));
            }
        }
    }
}
