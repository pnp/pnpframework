using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Utilities;
using System.Collections.Generic;

namespace PnP.Framework.Test.AppModelExtensions
{
    [TestClass]
    public class UrlUtilityTests
    {
        [TestMethod]
        public void ContainsInvalidCharsReturnsFalseForValidString()
        {
            string validString = "abd-123";
            Assert.IsFalse(validString.ContainsInvalidUrlChars());
        }

        [TestMethod]
        public void ContainsInvalidUrlCharsReturnsTrueForInvalidString()
        {
            var targetVals = new List<char> { '#', '%', '*', '\\', ':', '<', '>', '?', '/', '+', '|', '"' };

            targetVals.ForEach(v => Assert.IsTrue((string.Format("abc{0}abc", v).ContainsInvalidUrlChars())));
        }

        [TestMethod]
        public void StripInvalidUrlCharsReturnsStrippedString()
        {
            var invalidString = "a#%*\\:<>?/+|b";

            Assert.AreEqual("ab", invalidString.StripInvalidUrlChars());
        }

        [TestMethod]
        public void ReplaceInvalidUrlCharsReturnsStrippedString()
        {
            var invalidString = "a#%*\\:<>?/+|b";
            Assert.AreEqual("a---------------------------------b", invalidString.ReplaceInvalidUrlChars("---"));
        }

        [TestMethod]
        public void UrlPathEncodePerformsUrlEncodingButLeavesSlashesAlone()
        {
            var input = "/sites/site001/document library/folder abc";
            var expected = "/sites/site001/document%20library/folder%20abc";
            
            var actual = input.UrlPathEncode();

            Assert.AreEqual(expected, actual);
        }
    }
}
