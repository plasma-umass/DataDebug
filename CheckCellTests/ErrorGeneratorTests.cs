using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ErrorGenerator = UserSimulation.ErrorGenerator;
using UserSimulation;

namespace CheckCellTests
{
    [TestClass]
    class ErrorGeneratorTests
    {
        [TestMethod]
        public void TestErrorGenerator()
        {
            var eg = new ErrorGenerator();
            var result = eg.GenerateErrorString("blah");
        }
    }
}
