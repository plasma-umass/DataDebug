using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
        }
    }
}