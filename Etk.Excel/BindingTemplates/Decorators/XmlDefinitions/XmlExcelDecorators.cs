using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Etk.Tools.Extensions;

namespace Etk.Excel.BindingTemplates.Decorators.XmlDefinitions
{
    [XmlRoot("ExcelDecorators")]
    public class XmlExcelDecorators
    {
        [XmlElement(ElementName = "RangeDecorator", Type = typeof(XmlExcelRangeDecorator))]
        public List<XmlExcelRangeDecorator> RangeDecorators
        { get; set; }

        //[XmlElement(ElementName = "Decorator", Type = typeof(XmlExcelDecorator))]
        //public List<XmlExcelDecorator> Decorators
        //{ get; set; }

        public static XmlExcelDecorators CreateInstance(string definition)
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
                    throw new EtkException(string.Format("Cannot retrieve the Excel decorators from '{0}'. {1}", def, ex.Message));
                }
            }
            return ret;
        }
    }
}
