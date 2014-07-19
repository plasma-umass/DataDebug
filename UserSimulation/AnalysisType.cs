using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UserSimulation
{
    [Serializable]
    public enum AnalysisType
    {
        CheckCell5 = 0,    // 0.05
        CheckCell10 = 1,    // 0.10
        NormalPerRange = 2,    //normal analysis of inputs on a per-range granularity
        NormalAllInputs = 3,    //normal analysis on the entire set of inputs
        CheckCellN = 4     // CheckCell, n steps
    }
}
