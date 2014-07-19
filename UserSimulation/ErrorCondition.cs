using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UserSimulation
{
    [Serializable]
    public enum ErrorCondition
    {
        OK,
        ContainsNoInputs,
        Exception
    }
}
