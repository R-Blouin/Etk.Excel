namespace Etk.BindingTemplates.Definitions.Templates.Xml
{
    using System;
    using System.Xml.Serialization;

    [XmlRoot("Link")]
    public class XmlTemplateLink
    {
        [XmlAttribute]
        public string Name
        { get; set; }
        
        [XmlAttribute]
        public string Description
        { get; set; }
        
        [XmlAttribute]
        public string To
        { get; set; }

        [XmlAttribute]
        public string With
        { get; set; }

        [XmlIgnore]
        public LinkedTemplatePositioning Position
        { get; set; }

        private string positionStr;
        [XmlAttribute("Position")]
        public string PositionStr
        {
            get { return positionStr; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    positionStr = value.Trim().ToUpper();
                    if (positionStr.Equals("R") || positionStr.Equals("RELATIVE"))
                        Position = LinkedTemplatePositioning.Relative;
                    else if (positionStr.Equals("A") || positionStr.Equals("ABSOLUTE"))
                        Position = LinkedTemplatePositioning.Absolute;
                    else
                        throw new ArgumentException(string.Format("The attribut 'Position' '{0}' is invalid. Value must be 'Relative' (or 'R') or 'Absolute' (or 'A') (no case sentitive)", value));
                }
            }
        }
    }
}
