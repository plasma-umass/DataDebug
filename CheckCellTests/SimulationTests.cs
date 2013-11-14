using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;


namespace CheckCellTests
{
    [TestClass]
    public class SimulationTests
    {
        [TestMethod]
        public void SimulationRunTest()
        {
            //UserSimulation.Classification c = new UserSimulation.Classification();
            ////set typo dictionary to explicit one
            //Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();

            //var key = new Tuple<OptChar, string>(OptChar.Some('t'), "y");
            //typo_dict.Add(key, 1);

            //key = new Tuple<OptChar, string>(OptChar.Some('t'), "t");
            //typo_dict.Add(key, 0);

            //key = new Tuple<OptChar, string>(OptChar.Some('T'), "TT");
            //typo_dict.Add(key, 1);

            //key = new Tuple<OptChar, string>(OptChar.Some('e'), "e");
            //typo_dict.Add(key, 1);

            //key = new Tuple<OptChar, string>(OptChar.Some('s'), "s");
            //typo_dict.Add(key, 1);

            ////The transpositions dictionary is empty so no transpositions should occur
            //c.SetTypoDict(typo_dict);
            //c.Serialize();

            UserSimulation.Simulation sim = new UserSimulation.Simulation();
            sim.Run(100, "C:\\\\Users\\Dimitar Gochev\\Documents\\GitHub\\papers\\DataDebug\\OOPSLA-2013\\Spreadsheets\\test\\accurate_runtimes2.xlsx", 0.95, new Excel.Application(), 0.05);
        }
    }
}
