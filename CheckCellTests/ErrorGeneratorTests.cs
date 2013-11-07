using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ErrorGenerator = UserSimulation.ErrorGenerator;
using UserSimulation;
using Sign = LongestCommonSubsequence.Sign;

namespace CheckCellTests
{
    [TestClass]
    class ErrorGeneratorTests
    {
        [TestMethod]
        public void TestErrorGenerator()
        {
            var eg = new ErrorGenerator();

            //set dictionaries to explicit ones
            Dictionary<Tuple<Sign, Sign>, int> sign_dict = new Dictionary<Tuple<Sign, Sign>, int>();
            Dictionary<Tuple<char, string>, int> typo_dict = new Dictionary<Tuple<char, string>, int>();

           // sign_dict.Add(

            var result = eg.GenerateErrorString("Testing, testing, 123...");
        }
    }
}
