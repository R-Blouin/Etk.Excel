namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Xml.Serialization;

    public class Order
    {
        [XmlAttribute]
        public int Id
        { get; set; }

        [XmlIgnore]
        public int CustomerId
        { get; set; }

        [XmlAttribute("Date")]
        public string DateFomXml
        {
            get { return Date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture); }
            set { Date = DateTime.ParseExact(value, "yyyy/MM/dd", CultureInfo.InvariantCulture); }
        }

        [XmlElement(ElementName = "OrderLine", Type = typeof(OrderLine))]
        //[XmlElement(ElementName = "OrderLineWithDiscount", Type = typeof(OrderLineWithDiscount))]
        public List<OrderLine> Lines
        { get; set; }

        [XmlIgnore]
        public DateTime Date
        { get; private set; }
    }
}