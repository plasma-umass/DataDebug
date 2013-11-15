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
            UserSimulation.Simulation sim = new UserSimulation.Simulation();
            //sim.Run(100, "C:\\\\Users\\Dimitar Gochev\\Documents\\GitHub\\papers\\DataDebug\\OOPSLA-2013\\Spreadsheets\\test\\accurate_runtimes2.xlsx", 0.95, new Excel.Application(), 0.05);
            sim.Run(100, "C:\\\\Users\\Dimitar Gochev\\Documents\\GitHub\\papers\\DataDebug\\OOPSLA-2013\\Spreadsheets\\test\\Test2.xlsx", 0.95, new Excel.Application(), 0.05);
        }
    }
}
