namespace Etk.Excel.BindingTemplates.Controls.Button
{
    using System.Xml.Serialization;
    
    [XmlRoot("Button")]
    public class ExcelButtonDefinition
    {
        [XmlAttribute]
        public string Label
        { get; set; }

        [XmlAttribute]
        public int X
        { get; set; }

        [XmlAttribute]
        public int Y
        { get; set; }

        [XmlAttribute]
        public int W
        { get; set; }

        [XmlAttribute]
        public int H
        { get; set; }

        [XmlAttribute]
        public string Command
        { get; set; }

        [XmlAttribute]
        public string EnableProp
        { get; set; }

        public ExcelButtonDefinition()
        {
            X = Y = 0;
        }
    }
}
