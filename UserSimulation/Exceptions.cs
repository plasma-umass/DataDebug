using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UserSimulation
{
    public class NoRangeInputs : Exception { }
    public class NoFormulas : Exception { }
    public class SimulationNotRunException : Exception
    {
        public SimulationNotRunException(string message) : base(message) { }
    } 
}
