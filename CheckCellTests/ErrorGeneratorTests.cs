using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using ErrorGenerator = UserSimulation.ErrorGenerator;
using Classification = UserSimulation.Classification;
using Sign = LongestCommonSubsequence.Sign;

namespace CheckCellTests
{
    [TestClass]
    public class ErrorGeneratorTests
    {
        [TestMethod]
        public void TestErrorGenerator()
        {
            var eg = new ErrorGenerator();
            var c = new Classification();

            ////set dictionaries to explicit ones
            //Dictionary<Tuple<Sign, Sign>, int> sign_dict = new Dictionary<Tuple<Sign, Sign>, int>();
            //Dictionary<Tuple<char, string>, int> typo_dict = new Dictionary<Tuple<char, string>, int>();

            //var key = new Tuple<Sign, Sign>(Sign.Empty, Sign.Plus);

            //sign_dict.Add(key, 100);

            //key = new Tuple<Sign, Sign>(Sign.Empty, Sign.Minus);
            //sign_dict.Add(key, 100);

            //key = new Tuple<Sign, Sign>(Sign.Empty, Sign.Empty);
            //sign_dict.Add(key, 0);

            //key = new Tuple<Sign, Sign>(Sign.Minus, Sign.Empty);
            //sign_dict.Add(key, 100);

            //key = new Tuple<Sign, Sign>(Sign.Minus, Sign.Plus);
            //sign_dict.Add(key, 100);

            //key = new Tuple<Sign, Sign>(Sign.Minus, Sign.Minus);
            //sign_dict.Add(key, 0);

            //c.SetSignDict(sign_dict);
            var result = eg.GenerateErrorString("Test");
            string s = result.ToString();
        }
    }
}
