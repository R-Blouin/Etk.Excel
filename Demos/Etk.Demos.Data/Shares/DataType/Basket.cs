using System;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shares.DataType
{
    [Serializable]
    [XmlRoot("Basket")]
    public class Basket
    {
        [XmlAttribute]
        public int WaitingTime
        { get; set; }

        [XmlArray("Shares"), XmlArrayItem("Share")]
        public Share[] Shares
        { get; set; }

        [XmlArray("Forexes"), XmlArrayItem("Forex")]
        public Forex[] Forexes
        { get; set; }
    }
}
