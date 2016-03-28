using System.Xml.Serialization;

namespace Etk.Excel.BindingTemplates.Controls.CheckBox
{
    [XmlRoot("CheckBox")]
    public class ExcelCheckBoxDefinition
    {
        [XmlAttribute]
        public string Value
        { get; set; }

        //public string Label
        //{ get; set; }

        //[XmlAttribute]
        //public int X
        //{ get; set; }

        //[XmlAttribute]
        //public int Y
        //{ get; set; }

        //[XmlAttribute]
        //public int W
        //{ get; set; }

        //[XmlAttribute]
        //public int H
        //{ get; set; }

        public ExcelCheckBoxDefinition()
        {
            //X = Y = 0;
        }
    }
}
