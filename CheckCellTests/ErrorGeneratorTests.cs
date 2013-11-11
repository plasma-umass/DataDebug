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
        public void TestTypoGenerator()
        {
            var eg = new ErrorGenerator();
            var classification = new Classification();

            //set typo dictionary to explicit one
            Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();

            var key = new Tuple<OptChar, string>(OptChar.Some('t'), "y");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('t'), "t");
            typo_dict.Add(key, 0);

            key = new Tuple<OptChar, string>(OptChar.Some('T'), "TT");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('e'), "e");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('s'), "s");
            typo_dict.Add(key, 1);

            //The transpositions dictionary is empty so no transpositions should occur
            classification.SetTypoDict(typo_dict);
            var result = eg.GenerateErrorString("Testing", classification);
            string s = result.Item1;
            Assert.AreEqual("TTesying", s);
        }

        [TestMethod]
        public void TestTranspositionGenerator()
        {
            var eg = new ErrorGenerator();
            var classification = new Classification();

            //set typo dictionary to explicit one -- it's empty so no typos are possible
            Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();

            //Set the transpositions dictionary to explicit one
            Dictionary<int, int> transpositions_dict = new Dictionary<int, int>();
            transpositions_dict.Add(3, 10);
            //transpositions_dict.Add(0, 1);

            classification.SetTranspositionDict(transpositions_dict);
            classification.SetTypoDict(typo_dict);
            var result = eg.GenerateErrorString("abcd", classification);
            string s = result.Item1;
            Assert.AreEqual("abcd", s);
            
            //NOTE: Need a new ErrorGenerator for each test because the distribution tables are associated with it
            var eg2 = new ErrorGenerator();
            //Set the transpositions dictionary to explicit one
            var transpositions_dict2 = new Dictionary<int, int>();
            transpositions_dict2.Add(1, 10);
            //transpositions_dict2.Add(0, 1);
            Classification classification2 = new Classification();
            classification2.SetTranspositionDict(transpositions_dict2);
            classification2.SetTypoDict(typo_dict);
            var result2 = eg2.GenerateErrorString("abcd", classification2);
            string s2 = result2.Item1;
            Assert.AreEqual("abcd", s2);
            
            var eg3 = new ErrorGenerator();
            //Set the transpositions dictionary to explicit one
            var transpositions_dict3 = new Dictionary<int, int>();
            transpositions_dict3.Add(10, 10);
            transpositions_dict3.Add(-10, 10);
            transpositions_dict3.Add(0, 1);
            Classification classification3 = new Classification();
            classification3.SetTranspositionDict(transpositions_dict3);
            classification3.SetTypoDict(typo_dict);
            var result3 = eg3.GenerateErrorString("abcd", classification3);
            string s3 = result3.Item1;
            Assert.AreEqual("abcd", s3);
            Assert.AreEqual(0, result3.Item2.Count);
        }
    }
}