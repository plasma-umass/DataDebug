using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExtensionMethods
{
    public static class DataDebugExtensions
    {
        public static int ArgMax<T>(this IEnumerable<T> ie) where T : IComparable
        {
            var arr = ie.ToArray<T>();
            int argmax = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i].CompareTo(arr[argmax]) > 0)
                {
                    argmax = i;
                }
            }
            return argmax;
        }

    }
}
