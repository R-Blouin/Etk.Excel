using System.Xml.Serialization;

namespace Etk.Excel.BindingTemplates.Decorators.XmlDefinitions
{
    public class XmlExcelRangeDecorator
    {
        [XmlAttribute]
        public string Ident
        { get; set; }

        [XmlAttribute]
        public string Description
        { get; set; }

        [XmlAttribute]
        public string Method
        { get; set; }
        
        [XmlAttribute]
        public string Range
        { get; set; }

        [XmlAttribute]
        public bool NotOnlyColor
        { get; set; }
    }
}
