using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops.DataType
{
    [Serializable]
    [XmlRootAttribute("Products")]
    public class ProductList
    {
        [XmlElement("Product")]
        public List<Product> Products
        { get; set; }
    }
}
