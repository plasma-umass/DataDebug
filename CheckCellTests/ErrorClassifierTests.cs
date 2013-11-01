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

        public void TestDecimalOmission()
        {
            var c = new Classification();

            Assert.AreEqual(OptString.Some("1.2"), c.TestDecimalOmission("1.2","12"));
            Assert.AreEqual(OptString.None, c.TestDecimalOmission("12", "123"));
            Assert.AreEqual(OptString.None, c.TestDecimalOmission("12", "12"));
            Assert.AreEqual(OptString.Some("1.25"), c.TestDecimalOmission("1.25", "12"));
            Assert.AreEqual(OptString.Some("1.2"), c.TestDecimalOmission("1.2", "123"));
            Assert.AreEqual(OptString.Some("12345.6"), c.TestDecimalOmission("12345.6", "12"));
        }

        public void TestDecimalMisplacement()
        {
            var c = new Classification();

            Assert.AreEqual(OptString.Some("1.2"), c.TestMisplacedDecimal("1.2", "12."));
            Assert.AreEqual(OptString.None, c.TestMisplacedDecimal("12", "123"));
            Assert.AreEqual(OptString.Some("1.23"), c.TestMisplacedDecimal("1.2", ".123"));
            Assert.AreEqual(OptString.None, c.TestMisplacedDecimal("12", "1.23"));
            Assert.AreEqual(OptString.Some("1.20"), c.TestMisplacedDecimal("1.2345", "120."));         
        }
    }
}
