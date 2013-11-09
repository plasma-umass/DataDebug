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

            //set typo dictionary to explicit one
            Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();

            var key = new Tuple<OptChar, string>(OptChar.Some('t'), "t");
            typo_dict.Add(key, 0);

            key = new Tuple<OptChar, string>(OptChar.Some('t'), "blah");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('T'), "T--I-can't-type--T");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('e'), "e");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('s'), "s");
            typo_dict.Add(key, 1);

            //Set the transpositions dictionary to explicit one
            Dictionary<int, int> transpositions_dict = new Dictionary<int, int>();
            transpositions_dict.Add(1, 1);
            transpositions_dict.Add(2, 1);
            transpositions_dict.Add(0, 2);
            transpositions_dict.Add(-1, 1);
            
            classification.SetTranspositionDict(transpositions_dict);
            classification.SetTypoDict(typo_dict);
            var result = eg.GenerateErrorString("Testing", classification);
            string s = result.ToString();
        }
    }
}
