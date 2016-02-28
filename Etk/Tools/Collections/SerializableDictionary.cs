// From http://weblogs.asp.net/pwelter34/archive/2006/05/03/444961.aspx
///////////////////////////////////////////////////////////////////////

namespace Etk.Excel.UI.Collections
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Runtime.Serialization;
    using System.Xml;
    using System.Xml.Schema;
    using System.Xml.Serialization;

    [Serializable]
    public class SerializableDictionary<TKey, TVal> : Dictionary<TKey, TVal>, IXmlSerializable, ISerializable
    {
        #region Privateattributes and Properties
        private XmlSerializer valueSerializer;
        protected XmlSerializer ValueSerializer
        {
            get { return valueSerializer ?? (valueSerializer = new XmlSerializer(typeof(TVal))); }
        }

        private XmlSerializer keySerializer;
        private XmlSerializer KeySerializer
        {
            get { return keySerializer ?? (keySerializer = new XmlSerializer(typeof(TKey))); }
        }
        #endregion

        #region Constructors
        public SerializableDictionary()
        { }

        public SerializableDictionary(IDictionary<TKey, TVal> dictionary)
            : base(dictionary)
        { }

        public SerializableDictionary(IEqualityComparer<TKey> comparer)
            : base(comparer)
        { }

        public SerializableDictionary(int capacity)
            : base(capacity)
        { }

        public SerializableDictionary(IDictionary<TKey, TVal> dictionary, IEqualityComparer<TKey> comparer)
            : base(dictionary, comparer)
        { }

        public SerializableDictionary(int capacity, IEqualityComparer<TKey> comparer)
            : base(capacity, comparer)
        { }
        #endregion

        #region ISerializable Members
        protected SerializableDictionary(SerializationInfo info, StreamingContext context)
        {
            int itemCount = info.GetInt32("itemsCount");
            for (int i = 0; i < itemCount; i++)
            {
                KeyValuePair<TKey, TVal> kvp = (KeyValuePair<TKey, TVal>)info.GetValue(String.Format(CultureInfo.InvariantCulture, "Item{0}", i), typeof(KeyValuePair<TKey, TVal>));
                Add(kvp.Key, kvp.Value);
            }
        }

        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("itemsCount", Count);
            int itemIdx = 0;
            foreach (KeyValuePair<TKey, TVal> kvp in this)
            {
                info.AddValue(String.Format(CultureInfo.InvariantCulture, "Item{0}", itemIdx), kvp, typeof(KeyValuePair<TKey, TVal>));
                itemIdx++;
            }
        }
        #endregion

        #region IXmlSerializable Members
        void IXmlSerializable.WriteXml(XmlWriter writer)
        {
            foreach (KeyValuePair<TKey, TVal> kvp in this)
            {
                writer.WriteStartElement("item");
                writer.WriteStartElement("key");
                KeySerializer.Serialize(writer, kvp.Key);
                writer.WriteEndElement();
                writer.WriteStartElement("value");
                ValueSerializer.Serialize(writer, kvp.Value);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }

        void IXmlSerializable.ReadXml(XmlReader reader)
        {
            if (reader.IsEmptyElement)
                return;

            // Move past container
            if (reader.NodeType == XmlNodeType.Element && !reader.Read())
                throw new EtkException(string.Format("Error in Deserialization of SerializableDictionary<{0}, {1}>", typeof(TKey).FullName, typeof(TVal).FullName));
            while (reader.NodeType != XmlNodeType.EndElement)
            {
                reader.ReadStartElement("item");
                reader.ReadStartElement("key");
                TKey key = (TKey)KeySerializer.Deserialize(reader);
                reader.ReadEndElement();
                reader.ReadStartElement("value");
                TVal value = (TVal)ValueSerializer.Deserialize(reader);
                reader.ReadEndElement();
                reader.ReadEndElement();
                Add(key, value);
                reader.MoveToContent();
            }

            // Move past container
            if (reader.NodeType == XmlNodeType.EndElement)
                reader.ReadEndElement();
            else
                throw new EtkException(string.Format("Error in Deserialization of SerializableDictionary<{0}, {1}>", typeof(TKey).FullName, typeof(TVal).FullName));
        }

        XmlSchema IXmlSerializable.GetSchema()
        {
            return null;
        }
        #endregion
    }
}
