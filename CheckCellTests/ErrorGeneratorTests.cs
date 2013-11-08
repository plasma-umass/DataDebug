using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using ErrorGenerator = UserSimulation.ErrorGenerator;
using Classification = UserSimulation.Classification;
using Sign = LongestCommonSubsequence.Sign;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;


namespace CheckCellTests
{
    [TestClass]
    public class ErrorGeneratorTests
    {
        [TestMethod]
        public void TestErrorGenerator()
        {
            var eg = new ErrorGenerator();
            var classification = new Classification();

            ////set sign dictionary to explicit one
            //Dictionary<Tuple<Sign, Sign>, int> sign_dict = new Dictionary<Tuple<Sign, Sign>, int>();

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

            //set typo dictionary to explicit one
            Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();

            var key2 = new Tuple<OptChar, string>(OptChar.Some('t'), "t");
            typo_dict.Add(key2, 0);

            key2 = new Tuple<OptChar, string>(OptChar.Some('t'), "blah");
            typo_dict.Add(key2, 1);

            key2 = new Tuple<OptChar, string>(OptChar.Some('T'), "T--I-can't-type--T");
            typo_dict.Add(key2, 1);

            key2 = new Tuple<OptChar, string>(OptChar.Some('e'), "e");
            typo_dict.Add(key2, 1);

            key2 = new Tuple<OptChar, string>(OptChar.Some('s'), "s");
            typo_dict.Add(key2, 1);


            //classification.SetSignDict(sign_dict);
            classification.SetTypoDict(typo_dict);
            var result = eg.GenerateErrorString("Testing", classification);
            string s = result.ToString();
        }
    }
}
