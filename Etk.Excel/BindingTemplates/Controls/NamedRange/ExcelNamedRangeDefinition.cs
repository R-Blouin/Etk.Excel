using System.Xml.Serialization;

namespace Etk.Excel.BindingTemplates.Controls.NamedRange
{
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
