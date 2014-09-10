using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Numerics;

namespace DataDebugMethods
{
    public class FunctionOutput<T>
    {
        private T _value;
        private BigInteger _excludes;

        public FunctionOutput(T value, BigInteger excludes)
        {
            _value = value;
            _excludes = excludes;
        }

        public T GetValue()
        {
            return _value;
        }

        public BigInteger GetExcludes()
        {
            return _excludes;
        }

        public HashSet<int> GetExcludesAsHashSet()
        {
            var bits = new HashSet<int>();
            var e = _excludes;
            int i = 0;
            while (e != BigInteger.Zero)
            {
                // if the LSB is set, then the
                // ith index is excluded
                if ((BigInteger.One & e) == BigInteger.One)
                {
                    bits.Add(i);
                }
                // right shift
                e = e >> 1;
                // increment counter
                i += 1;
            }
            return bits;
        }
    }
}
