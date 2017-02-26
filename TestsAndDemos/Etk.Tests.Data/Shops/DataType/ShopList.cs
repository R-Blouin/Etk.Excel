namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;

    [Serializable]
    [XmlRootAttribute("Shops")]
    public class ShopList
    {
        [XmlElement("Shop")]
        public List<Shop> Shops
        { get; set; }
    }
}
