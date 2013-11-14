using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;


namespace CheckCellTests
{
    [TestClass]
    public class SimulationTests
    {
        [TestMethod]
        public void SimulationRunTest()
        {
            UserSimulation.Simulation sim = new UserSimulation.Simulation();
            sim.Run(100, "C:\\\\Users\\Dimitar Gochev\\Documents\\GitHub\\papers\\DataDebug\\OOPSLA-2013\\Spreadsheets\\test\\accurate_runtimes.xlsx", 0.95, new Excel.Application(), 0.05);
        }
    }
}
