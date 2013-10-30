using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataDebugMethods;
using OptString = Microsoft.FSharp.Core.FSharpOption<string>;

namespace CheckCellTests
{
    [TestClass]
    public class ErrorClassifierTests
    {
        [TestMethod]
        public void TestHasSignError()
        {
            var c = new Classification();

            Assert.AreEqual(OptString.Some("-3748"), c.HasSignError("-3748", "3748"));
            Assert.AreEqual(OptString.Some("+3748"), c.HasSignError("+3748", "3748"));
            Assert.AreEqual(OptString.Some("3748"), c.HasSignError("3748", "-3748"));
            Assert.AreEqual(OptString.Some("3748"), c.HasSignError("3748", "+3748"));
            Assert.AreEqual(OptString.None, c.HasSignError("3748", "3748"));
            Assert.AreEqual(OptString.None, c.HasSignError("3748", "33748"));
            Assert.AreEqual(OptString.None, c.HasSignError("3748", ""));
            Assert.AreEqual(OptString.None, c.HasSignError("-3748", "-g3748"));
            Assert.AreEqual(OptString.None, c.HasSignError("+3748", "+3748"));
        }
    }
}
