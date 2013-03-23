using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class InputSample
    {
        private int _i = 0;             // internal length counter for Add
        private string[] _input_array;  // the actual values of this array
        private HashSet<int> _excludes; // list of inputs excluded in this sample
        private int[] _includes;        // a counter of values included by this sample

        public InputSample(int size)
        {
            _input_array = new string[size];
            _excludes = new HashSet<int>();
        }
        public void Add(string value)
        {
            Debug.Assert(_i < _input_array.Length);
            _input_array[_i] = value;
            _i++;
        }
        public string GetInput(int num)
        {
            Debug.Assert(num < _input_array.Length);
            return _input_array[num];
        }
        public int Length()
        {
            return _input_array.Length;
        }
        public HashSet<int> GetExcludes()
        {
            return _excludes;
        }
        public int[] GetIncludes()
        {
            return _includes;
        }
        public void SetIncludes(int[] includes)
        {
            _includes = includes;
            for (int i = 0; i < includes.Length; i++)
            {
                if (includes[i] == 0)
                {
                    _excludes.Add(i);
                }
            }
        }
        public override int GetHashCode()
        {
            // note that in C#, shift never causes overflow
            int sum = 0;
            for (int i = 0; i < _includes.Length; i++)
            {
                sum += _includes[i] << i;
            }
            return sum;
        }
        public override bool Equals(object obj)
        {
            InputSample other = (InputSample)obj;
            return _includes.SequenceEqual(other.GetIncludes());
        }
        public override string ToString()
        {
            return String.Join(",", _input_array);
        }
    }
}
