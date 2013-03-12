using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebugMethods
{
    public class FunctionOutput
    {
        private string _value;
        private HashSet<int> _excludes;
        public FunctionOutput(string value, HashSet<int> excludes)
        {
            _value = value;
            _excludes = excludes;
        }
    }
}
