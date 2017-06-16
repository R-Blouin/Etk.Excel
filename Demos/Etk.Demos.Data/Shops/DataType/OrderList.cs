using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops.DataType
{
    [Serializable]
    [XmlRootAttribute("Orders")]
    public class OrderList
    {
        [XmlElement("Order")]
        public List<Order> Orders
        { get; set; }
    }
}
