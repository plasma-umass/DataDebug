using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebugMethods
{
    public class FunctionOutput<T>
    {
        private T _value;
        private HashSet<int> _excludes;

        public FunctionOutput(T value, HashSet<int> excludes)
        {
            _value = value;
            _excludes = excludes;
        }

        public T GetValue()
        {
            return _value;
        }

        public HashSet<int> GetExcludes()
        {
            return _excludes;
        }
    }
}
