using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;
using FSIntList = Microsoft.FSharp.Collections.FSharpList<int>;

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
        public void FixTranspositionTest()
        {
            var orig = "aaabq";
            var entered = "bbzzzaaaq";

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

            Assert.AreEqual(-3, deltas.Head);
            Assert.AreEqual("bzzzaaabq", entered2);
            int[] correct_additions = { 0, 1, 2, 3 };
            Assert.AreEqual(true, correct_additions.SequenceEqual<int>(additions2));
            int[] correct_omissions = {};
            Assert.AreEqual(true, correct_omissions.SequenceEqual<int>(omissions2));
            Tuple<int,int>[] correct_alignments = { new Tuple<int,int>(0,4), new Tuple<int,int>(1,5), new Tuple<int,int>(2,6), new Tuple<int,int>(3,7), new Tuple<int,int>(4,8) };
            Assert.AreEqual(true, correct_alignments.SequenceEqual<Tuple<int,int>>(alignments2));
        }

        [TestMethod]
        public void FixTranspositionTest2()
        {
            var orig = "baaaq";
            var entered = "zzzaaabbq";

            var alignments = LongestCommonSubsequence.LeftAlignedLCS(orig, entered);
            var additions = LongestCommonSubsequence.GetAddedCharIndices(entered, alignments);
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(orig, alignments);
            var fixedouts = LongestCommonSubsequence.FixTranspositions(alignments, additions, omissions, orig, entered);

            var entered2 = fixedouts.Item1;
            var alignments2 = fixedouts.Item2;
            var additions2 = fixedouts.Item3;
            var omissions2 = fixedouts.Item4;
            var deltas = fixedouts.Item5;

            Assert.AreEqual(3, deltas.Head);
            Assert.AreEqual("bzzzaaabq", entered2);
            int[] correct_additions = { 1, 2, 3, 7 };
            Assert.AreEqual(true, correct_additions.SequenceEqual<int>(additions2));
            int[] correct_omissions = { };
            Assert.AreEqual(true, correct_omissions.SequenceEqual<int>(omissions2));
            Tuple<int, int>[] correct_alignments = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 4), new Tuple<int, int>(2, 5), new Tuple<int, int>(3, 6), new Tuple<int, int>(4, 8) };
            Assert.AreEqual(true, correct_alignments.SequenceEqual<Tuple<int, int>>(alignments2));
        }

        [TestMethod]
        public void EnumFixTranspositionTests()
        {
            var orig1 = "acc";
            var ent1 = "cca";
            var al1 = LongestCommonSubsequence.LeftAlignedLCS(orig1, ent1);
            var ad1 = LongestCommonSubsequence.GetAddedCharIndices(ent1, al1);
            var om1 = LongestCommonSubsequence.GetMissingCharIndices(orig1, al1);
            var fix1 = LongestCommonSubsequence.FixTranspositions(al1, ad1, om1, orig1, ent1);
            Assert.AreEqual(orig1, fix1.Item1);
            Tuple<int, int>[] correct_alignments = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 2) };
            Assert.AreEqual(true, correct_alignments.SequenceEqual<Tuple<int, int>>(fix1.Item2));
            int [] correct_additions = {};
            Assert.AreEqual(true, correct_additions.SequenceEqual<int>(fix1.Item3));
            int[] correct_omissions = { };
            Assert.AreEqual(true, correct_omissions.SequenceEqual<int>(fix1.Item4));
            Assert.AreEqual(2, fix1.Item5.Head);

            var orig2 = "acc";
            var ent2 = "cac";
            // this line is to avoid nondeterministic choice of alignments
            Tuple<int, int>[] al2_a = { new Tuple<int, int>(0, 1), new Tuple<int, int>(2, 2) };
            var al2 = LongestCommonSubsequence.ToFSList(al2_a);
            var ad2 = LongestCommonSubsequence.GetAddedCharIndices(ent2, al2);
            var om2 = LongestCommonSubsequence.GetMissingCharIndices(orig2, al2);
            var fix2 = LongestCommonSubsequence.FixTranspositions(al2, ad2, om2, orig2, ent2);
            Assert.AreEqual(orig2, fix2.Item1);
            Tuple<int, int>[] correct_alignments2 = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 2) };
            Assert.AreEqual(true, correct_alignments2.SequenceEqual<Tuple<int, int>>(fix2.Item2));
            int[] correct_additions2 = { };
            Assert.AreEqual(true, correct_additions2.SequenceEqual<int>(fix2.Item3));
            int[] correct_omissions2 = { };
            Assert.AreEqual(true, correct_omissions2.SequenceEqual<int>(fix2.Item4));
            Assert.AreEqual(-1, fix2.Item5.Head);

            var orig3 = "cac";
            var ent3 = "acc";
            // this line is to avoid nondeterministic choice of alignments
            Tuple<int, int>[] al3_a = { new Tuple<int, int>(0, 1), new Tuple<int, int>(2, 2) };
            var al3 = LongestCommonSubsequence.ToFSList(al3_a);
            var ad3 = LongestCommonSubsequence.GetAddedCharIndices(ent3, al3);
            var om3 = LongestCommonSubsequence.GetMissingCharIndices(orig3, al3);
            var fix3 = LongestCommonSubsequence.FixTranspositions(al3, ad3, om3, orig3, ent3);
            Assert.AreEqual(orig3, fix3.Item1);
            Tuple<int, int>[] correct_alignments3 = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 1), new Tuple<int, int>(2, 2) };
            Assert.AreEqual(true, correct_alignments3.SequenceEqual<Tuple<int, int>>(fix2.Item2));
            int[] correct_additions3 = { };
            Assert.AreEqual(true, correct_additions3.SequenceEqual<int>(fix2.Item3));
            int[] correct_omissions3 = { };
            Assert.AreEqual(true, correct_omissions3.SequenceEqual<int>(fix2.Item4));
            Assert.AreEqual(-1, fix3.Item5.Head);
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