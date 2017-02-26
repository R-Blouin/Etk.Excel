namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;

    [Serializable]
    [XmlRootAttribute("Customers")]
    public class CustomerList
    {
        [XmlElement("Customer")]
        public List<Customer> Customers
        { get; set; }
    }
}
