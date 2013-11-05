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
            Tuple<int, int>[] shouldnotbe_a = { new Tuple<int, int>(0, 0), new Tuple<int, int>(1, 2), new Tuple<int, int>(2, 4), new Tuple<int, int>(3, 5), new Tuple<int, int>(4, 6) };
            var shouldbe = new System.Collections.Generic.List<Tuple<int, int>>(shouldbe_a);
            var shouldnotbe = new System.Collections.Generic.List<Tuple<int, int>>(shouldnotbe_a);
            var sss = LongestCommonSubsequence.LCS_Hash_Char(s1, s2);
            foreach (var ss in sss)
            {
                Assert.AreEqual(true, ss.SequenceEqual<Tuple<int,int>>(shouldbe));
                Assert.AreNotEqual(true, ss.SequenceEqual<Tuple<int, int>>(shouldnotbe));
            }
        }
    }
}