namespace Etk.Excel.BindingTemplates.Controls.NamedRange
{
    using System.Xml.Serialization;
    
    [XmlRoot("NR")]
    public class ExcelNamedRangeDefinition
    {
        [XmlAttribute]
        public string Name
        { get; set; }

        [XmlText]
        public string Value 
        { get; set; }
    }
}
