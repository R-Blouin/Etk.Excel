// From http://www.blackwasp.co.uk/ReadOnlyDictionary.aspx
///////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Etk.Tools.Collections
{
    public sealed class ReadOnlyDictionary<TKey, TValue> : IDictionary<TKey, TValue>
    {
        private IDictionary<TKey, TValue> _dictionary;

        public ReadOnlyDictionary(IDictionary<TKey, TValue> source)
        {
            _dictionary = source;
        }

        public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator()
        {
            foreach (KeyValuePair<TKey, TValue> item in _dictionary)
            {
                yield return item;
            }
        }

        public bool ContainsKey(TKey key)
        {
            return _dictionary.ContainsKey(key);
        }


        public bool TryGetValue(TKey key, out TValue value)
        {
            return _dictionary.TryGetValue(key, out value);
        }

        public bool Contains(KeyValuePair<TKey, TValue> item)
        {
            return _dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex)
        {
            _dictionary.CopyTo(array, arrayIndex);
        }


        public TValue this[TKey key]
        {
            get { return _dictionary[key]; }
        }


        public ICollection<TKey> Keys
        {
            get
            {
                ReadOnlyCollection<TKey> keys = new ReadOnlyCollection<TKey>(new List<TKey>(_dictionary.Keys));
                return (ICollection<TKey>)keys;
            }
        }


        public ICollection<TValue> Values
        {
            get
            {
                ReadOnlyCollection<TValue> values = new ReadOnlyCollection<TValue>(new List<TValue>(_dictionary.Values));
                return (ICollection<TValue>)values;
            }
        }


        public int Count
        {
            get { return _dictionary.Count; }
        }


        public bool IsReadOnly
        {
            get { return true; }
        }


        void IDictionary<TKey, TValue>.Add(TKey key, TValue value)
        {
            throw new NotSupportedException();
        }


        bool IDictionary<TKey, TValue>.Remove(TKey key)
        {
            throw new NotSupportedException();
        }


        TValue IDictionary<TKey, TValue>.this[TKey key]
        {
            get { return this[key]; }
            set { throw new NotSupportedException(); }
        }


        void ICollection<KeyValuePair<TKey, TValue>>.Add(KeyValuePair<TKey, TValue> item)
        {
            throw new NotSupportedException();
        }


        bool ICollection<KeyValuePair<TKey, TValue>>.Remove(KeyValuePair<TKey, TValue> item)
        {
            throw new NotSupportedException();
        }


        void ICollection<KeyValuePair<TKey, TValue>>.Clear()
        {
            throw new NotSupportedException();
        }


        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
