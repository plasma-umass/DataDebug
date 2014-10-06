using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebugMethods
{
    public class BiDictionary<T,U>
    {
        Dictionary<T, U> _dict1 = new Dictionary<T, U>();
        Dictionary<U, T> _dict2 = new Dictionary<U, T>();

        public void Add(T v1, U v2)
        {
            _dict1.Add(v1, v2);
            _dict2.Add(v2, v1);
        }

        public bool ContainsKey(T key)
        {
            return _dict1.ContainsKey(key);
        }

        public bool ContainsKey(U key)
        {
            return _dict2.ContainsKey(key);
        }

        public bool Remove(T key)
        {
            return _dict2.Remove(_dict1[key]) &&
                   _dict1.Remove(key);
        }

        public bool Remove(U key)
        {
            return _dict1.Remove(_dict2[key]) &&
                   _dict2.Remove(key);
        }

        public bool TryGetValue(T key, out U value)
        {
            return _dict1.TryGetValue(key, out value);
        }

        public bool TryGetValue(U key, out T value)
        {
            return _dict2.TryGetValue(key, out value);
        }

        public U this[T key]
        {
            get
            {
                return _dict1[key];
            }
            set
            {
                U oldval = _dict1[key];
                _dict1[key] = value;
                _dict2.Remove(oldval);
                _dict2.Add(value, key);
            }
        }

        public T this[U key]
        {
            get
            {
                return _dict2[key];
            }
            set
            {
                T oldval = _dict2[key];
                _dict2[key] = value;
                _dict1.Remove(oldval);
                _dict1.Add(value, key);
            }
        }

        public void Add(System.Collections.Generic.KeyValuePair<T, U> item)
        {
            _dict1.Add(item.Key, item.Value);
            _dict2.Add(item.Value, item.Key);
        }

        public void Add(System.Collections.Generic.KeyValuePair<U, T> item)
        {
            _dict2.Add(item.Key, item.Value);
            _dict1.Add(item.Value, item.Key);
        }

        public void Clear()
        {
            _dict1.Clear();
            _dict2.Clear();
        }

        public bool Contains(System.Collections.Generic.KeyValuePair<T, U> item)
        {
            var vkp = new KeyValuePair<U, T>(item.Value, item.Key);
            return _dict1.Contains(item) && _dict2.Contains(vkp);
        }

        public int Count
        {
            get { return _dict1.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(System.Collections.Generic.KeyValuePair<T, U> item)
        {
            return _dict1.Remove(item.Key) && _dict2.Remove(item.Value);
        }

        public bool Remove(System.Collections.Generic.KeyValuePair<U, T> item)
        {
            return _dict1.Remove(item.Value) && _dict2.Remove(item.Key);
        }

        public IEnumerable<KeyValuePair<T, U>> AsTUEnum()
        {
            return _dict1;
        }

        public IEnumerable<KeyValuePair<U,T>> AsUTEnum()
        {
            return _dict2;
        }
    }
}
