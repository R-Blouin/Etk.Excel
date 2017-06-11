using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops.DataType
{
    [Serializable]
    [XmlRootAttribute("Shops")]
    public class ShopList
    {
        [XmlElement("Shop")]
        public List<Shop> Shops
        { get; set; }
    }
}
