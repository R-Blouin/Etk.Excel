namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;

    [Serializable]
    [XmlRootAttribute("Orders")]
    public class OrderList
    {
        [XmlElement("Order")]
        public List<Order> Orders
        { get; set; }
    }
}
