using System;
using System.Security;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Utilities;

namespace PnP.Framework.Tests.Utilities
{
    [TestClass]
    public class EncryptionUtilityTests
    {
        [TestMethod]
        public void ToSecureStringTest()
        {
            var testString = "this is a test string";
            var secureTestString = EncryptionUtility.ToSecureString(testString);
            Assert.IsInstanceOfType(secureTestString, typeof(SecureString));
            Assert.IsNotNull(testString);
            Assert.IsTrue(testString.Length > 0);
        }

        [TestMethod]
        public void ToInSecureStringTest()
        {
            var testString = "this is a test string";
            var secureTestString = EncryptionUtility.ToSecureString(testString);

            var insecureTestString = EncryptionUtility.ToInsecureString(secureTestString);

            Assert.IsInstanceOfType(insecureTestString, typeof(string));
            Assert.IsNotNull(insecureTestString);
            Assert.IsTrue(insecureTestString.Length > 0);
            Assert.IsTrue(testString == insecureTestString);
        }

    }
}
