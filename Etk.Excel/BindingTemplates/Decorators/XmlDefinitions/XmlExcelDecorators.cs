namespace Etk.Excel.BindingTemplates.Decorators.XmlDefinitions
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;
    using Etk.Excel.UI.Extensions;
    
    [XmlRoot("ExcelDecorators")]
    public class XmlExcelDecorators
    {
        [XmlElement(ElementName = "RangeDecorator", Type = typeof(XmlExcelRangeDecorator))]
        public List<XmlExcelRangeDecorator> RangeDecorators
        { get; set; }

        //[XmlElement(ElementName = "Decorator", Type = typeof(XmlExcelDecorator))]
        //public List<XmlExcelDecorator> Decorators
        //{ get; set; }

        static public XmlExcelDecorators CreateInstance(string definition)
        {
            XmlExcelDecorators ret = null;
            definition = definition.EmptyIfNull().Trim();
            if (!string.IsNullOrEmpty(definition))
            {
                try
                {
                    ret = definition.Deserialize<XmlExcelDecorators>();
                }
                catch (Exception ex)
                {
                    string def = definition.EmptyIfNull().Trim();
                    if (def.Length > 150)
                        def = def.Substring(0, 149) + "...";

                    string message = string.Format("Cannot retrieve the Excel decorators from '{0}'. {1}", def, ex.Message);
                    throw new EtkException(message, ex);
                }
            }
            return ret;
        }
    }
}
