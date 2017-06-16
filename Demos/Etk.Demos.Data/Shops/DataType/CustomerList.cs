using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops.DataType
{
    [Serializable]
    [XmlRootAttribute("Customers")]
    public class CustomerList
    {
        [XmlElement("Customer")]
        public List<Customer> Customers
        { get; set; }
    }
}
