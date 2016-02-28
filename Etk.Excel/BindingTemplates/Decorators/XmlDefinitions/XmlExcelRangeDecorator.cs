namespace Etk.Excel.BindingTemplates.Decorators.XmlDefinitions
{
    using System.Xml.Serialization;

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
