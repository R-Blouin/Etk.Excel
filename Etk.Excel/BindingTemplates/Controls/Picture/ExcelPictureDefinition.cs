namespace Etk.Excel.BindingTemplates.Controls.Picture
{
    using System.Xml.Serialization;
    
    [XmlRoot("Picture")]
    public class ExcelPictureDefinition
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

        public ExcelPictureDefinition()
        {
            //X = Y = 0;
        }
    }
}
