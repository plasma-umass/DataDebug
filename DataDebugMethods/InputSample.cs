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
        private object[,] _input_array; // this is stored in Excel ONE-BASED format, DANGER!!!
        private HashSet<int> _excludes; // list of inputs excluded in this sample
        private int[] _includes;        // a counter of values included by this sample
        private int _rows;
        private int _cols;

        public InputSample(int rows, int cols)
        {
            _rows = rows;
            _cols = cols;
            _excludes = new HashSet<int>();
        }

        public void Add(string datum) {
            if (_i == 0)
            {
                // initialize _input_array
                Int32[] lowerBounds = { 1, 1 };
                Int32[] lengths = { _rows, _cols };
                _input_array = (object[,])Array.CreateInstance(typeof(object), lengths, lowerBounds);
            }
            var pair = OneDToTwoD(_i);
            var col_idx = pair.Item1;
            var row_idx = pair.Item2;
            _input_array[col_idx, row_idx] = datum;
            _i++;
        }

        public void AddArray(object[,] data)
        {
            if (_i != 0)
            {
                throw new Exception("You must use EITHER Add or AddArray, but not both.");
            }
            _input_array = data;
        }

        // this also adds a one to each index
        // Excel is also row-major
        private Tuple<int,int> OneDToTwoD(int idx) {
            var row_idx = (idx % _cols) + 1;
            var col_idx = (idx / _cols) + 1;
            return new Tuple<int,int>(col_idx, row_idx);
        }

        public string GetInput(int num)
        {
            // we assign a numbering scheme from
            // topleft to bottom right, starting at 0
            if (num <= _input_array.Length)
            {
                throw new Exception("num <= _input_array.Length");
            }
            var pair = OneDToTwoD(num);
            var col_idx = pair.Item1;
            var row_idx = pair.Item2;
            return System.Convert.ToString(_input_array[col_idx, row_idx]);
        }

        public int Length()
        {
            return _input_array.Length;
        }
        
        public int Rows() 
        { 
            return _rows; 
        }
        
        public int Columns() 
        { 
            return _cols; 
        }
        
        public HashSet<int> GetExcludes()
        {
            return _excludes;
        }

        public int[] GetIncludes()
        {
            return _includes;
        }

        // indicate which values are included in a resample
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

        public object[,] GetInputArray()
        {
            //// create one-based 2D multi-array
            //var output = Array.CreateInstance(typeof(object), new int[2] { _input_array.GetLength(0), _input_array.GetLength(1) }, new int[2] { 1, 1});

            return _input_array;
        }

        public string Text()
        {
            string text = "";
            foreach (object obj in _input_array)
            {
                text += obj + ",";
            }
            text.Remove(text.LastIndexOf(',') - 1);
            return text;
        }
    }
}
