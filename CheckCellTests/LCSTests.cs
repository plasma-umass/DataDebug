using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;

namespace CheckCellTests
{
    [TestClass]
    public class LCSTests
    {
        [TestMethod]
        public void TestSubstring()
        {
            var s1 = "Hello";
            var s2 = "Helloo";
            var ss = LongestCommonSubsequence.LCS_Hash(s1, s2);
            Assert.AreEqual(true, ss.Contains(s1));
        }

        [TestMethod]
        public void TestCharSequence()
        {
            var s1 = "Hello";
            var s2 = "Heellloo";
            Tuple<int, int>[] shouldbe_a = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 2), new Tuple<int, int>(2, 4), new Tuple<int, int>(3, 5), new Tuple<int, int>(4, 7) };
            Tuple<int, int>[] shouldalsobe_a = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 2), new Tuple<int, int>(2, 4), new Tuple<int, int>(3, 5), new Tuple<int, int>(4, 6) };
            var shouldbe = new System.Collections.Generic.List<Tuple<int, int>>(shouldbe_a);
            var shouldalsobe = new System.Collections.Generic.List<Tuple<int, int>>(shouldalsobe_a);
            var sss = LongestCommonSubsequence.LCS_Hash_Char(s1, s2);
            var found = 0;
            foreach (var ss in sss)
            {
                if (ss.SequenceEqual<Tuple<int,int>>(shouldbe) || ss.SequenceEqual<Tuple<int, int>>(shouldalsobe))
                {
                    found +=1 ;
                }
            }
            Assert.AreEqual<int>(2, found);
            Assert.AreEqual(12, sss.Count);
        }

        [TestMethod]
        public void TestCharSequences()
        {
            var s1 = "abc";
            var s2 = "abcc";
            Tuple<int, int>[] shouldbe_1_a = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 2) };
            Tuple<int, int>[] shouldbe_2_a = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 3) };
            var shouldbe_1 = new System.Collections.Generic.List<Tuple<int, int>>(shouldbe_1_a);
            var shouldbe_2 = new System.Collections.Generic.List<Tuple<int, int>>(shouldbe_2_a);
            var sss = LongestCommonSubsequence.LCS_Hash_Char(s1, s2);
            Assert.AreEqual(2, sss.Count);
        }

        [TestMethod]
        public void TestLeftAlignedLCS()
        {
            var s1 = "abc";
            var s2 = "abcc";
            // this returns exactly one left-aligned common subsequence, chosen randomly
            var ss = LongestCommonSubsequence.LeftAlignedLCSList(s1, s2);

            // but in this case, both left-aligned subsequences will be the same
            Tuple<int, int>[] shouldbe = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 2) };

            Assert.AreEqual(true, ss.SequenceEqual<Tuple<int, int>>(shouldbe));

            var s3 = "aaab";
            var s4 = "bzzzaaa";
            var ss2 = LongestCommonSubsequence.LeftAlignedLCS(s3, s4);

            Tuple<int, int>[] shouldbe2 = { new Tuple<int, int>(0, 4), new Tuple<int, int>(1, 5), new Tuple<int, int>(2, 6) };

            Assert.AreEqual(true, ss2.SequenceEqual<Tuple<int, int>>(shouldbe2));
        }

        [TestMethod]
        public void IndicesAppended()
        {
            var s1 = "abc";
            var s2 = "abcc";
            // this returns exactly one left-aligned common subsequence, chosen randomly
            var ss = LongestCommonSubsequence.LeftAlignedLCS(s1, s2);

            // get the appended indices
            var idxs = LongestCommonSubsequence.GetAddedCharIndices(s2, ss);

            Assert.AreEqual(1, idxs.Count());
            Assert.AreEqual(idxs[0], 3);
        }

        [TestMethod]
        public void IndicesExcluded()
        {
            var s1 = "abc";
            var s2 = "bc";
            // this returns exactly one left-aligned common subsequence, chosen randomly
            var ss = LongestCommonSubsequence.LeftAlignedLCS(s1, s2);

            // get the excluded indices
            var idxs = LongestCommonSubsequence.GetMissingCharIndices(s1, ss);

            Assert.AreEqual(1, idxs.Count());
            Assert.AreEqual(idxs[0], 0);
        }

        [TestMethod]
        public void TranspositionTest()
        {
            var s1 = "abc";
            var s2 = "acb";

            var ss = LongestCommonSubsequence.LeftAlignedLCS(s1, s2);
            var additions = LongestCommonSubsequence.GetAddedCharIndices(s2, ss);
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(s1, ss);
            var transpositions = LongestCommonSubsequence.GetTranspositions(additions, omissions, s1, s2, Microsoft.FSharp.Collections.FSharpList<Tuple<int,int>>.Empty);

            // there should be only 1 transposition here, but depending
            // on the randomly-chosen alignment, it may be one or the other
            var t1 = new Tuple<int,int>(2,1);
            var t2 = new Tuple<int,int>(1,2);
            Tuple<int, int> TR = transpositions[0];
            Assert.AreEqual(true, TR.Equals(t1) || TR.Equals(t2));
        }

        [TestMethod]
        public void NoTranspositionTest()
        {
            var s1 = "abc";
            var s2 = "abc";

            var ss = LongestCommonSubsequence.LeftAlignedLCS(s1, s2);
            var additions = LongestCommonSubsequence.GetAddedCharIndices(s2, ss);
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(s1, ss);
            var transpositions = LongestCommonSubsequence.GetTranspositions(additions, omissions, s1, s2, Microsoft.FSharp.Collections.FSharpList<Tuple<int, int>>.Empty);
            Assert.AreEqual(0, transpositions.Length);
        }

        [TestMethod]
        public void FixTranspositionTest()
        {
            var orig = "aaab";
            var entered = "bzzzaaa";

            var ta = LongestCommonSubsequence.LCS_Char(orig, entered);
            var alignments = LongestCommonSubsequence.LeftAlignedLCS(orig, entered);
            var additions = LongestCommonSubsequence.GetAddedCharIndices(entered, alignments);
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(orig, alignments);
            var fixedouts = LongestCommonSubsequence.FixTranspositions(alignments, additions, omissions, orig, entered);

            var entered2 = fixedouts.Item1;
            var alignments2 = fixedouts.Item2;
            var additions2 = fixedouts.Item3;
            var omissions2 = fixedouts.Item4;
            var deltas = fixedouts.Item5;
            var a = "hi";
        }

        [TestMethod]
        public void TypoTest()
        {
            var s1 = "abcd";
            var s2 = "zdcd";

            var ss = LongestCommonSubsequence.LeftAlignedLCS(s1, s2);
            var typos = LongestCommonSubsequence.GetTypos(ss, s1, s2);

            OptChar[] keys = { OptChar.None, new OptChar('a'), new OptChar('b'), new OptChar('c'), new OptChar('d') };
            char[] values = { 'z', 'd', 'c', 'd' };

            var key_hs = new System.Collections.Generic.HashSet<OptChar>(keys);
            var value_hs = new System.Collections.Generic.HashSet<char>(values);

            var keys_seen = new System.Collections.Generic.HashSet<OptChar>();
            var values_seen = new System.Collections.Generic.HashSet<char>(); ;
            foreach (var typo in typos)
            {
                var key = typo.Item1;
                var str = typo.Item2;
                keys_seen.Add(key);
                foreach (char c in str)
                {
                    values_seen.Add(c);
                }
            }

            Assert.AreEqual(true, key_hs.SetEquals(keys_seen));
            Assert.AreEqual(true, value_hs.SetEquals(values_seen));
        }
    }
}