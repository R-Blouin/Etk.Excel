namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;

    [Serializable]
    [XmlRootAttribute("Products")]
    public class ProductList
    {
        [XmlElement("Product")]
        public List<Product> Products
        { get; set; }
    }
}
