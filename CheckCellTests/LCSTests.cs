using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CheckCellTests
{
    [TestClass]
    public class LCSTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            var s1 = "Hello";
            var s2 = "Helloo";
            var ss = LongestCommonSubsequence.LCS_Hash(s1, s2);
            Assert.AreEqual(true, ss.Contains(s1));
        }
    }
}
