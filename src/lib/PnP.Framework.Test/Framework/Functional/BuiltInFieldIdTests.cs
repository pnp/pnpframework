using System;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Enums;

namespace PnP.Framework.Test.Framework.Functional
{
    /// <summary>
    /// Test cases for checking the functionality of BuiltInFieldId
    /// </summary>
    [TestClass]
    public class BuiltInFieldIdTests 
    {
        [TestMethod]
        public void KnownBuiltInFieldCanBeFound()
        {
            var knownExistingId = BuiltInFieldId.ContentTypeId;

            var actual = BuiltInFieldId.Contains(knownExistingId);
            
            Assert.AreEqual(true, actual);
        }
        
        [TestMethod]
        public void KnownNotBuiltInFieldCanNotBeFound()
        {
            var knownNonExistingId = new Guid("11111111-1111-1111-1111-111111111111");

            var actual = BuiltInFieldId.Contains(knownNonExistingId);
            
            Assert.AreEqual(false, actual);
        }
        
        [TestMethod]
        public void AllKnownFieldIdsCanBeFoundInTheContainsMethod()
        {
            var type = typeof(BuiltInFieldId);
            var allIdFields = type
                .GetFields(BindingFlags.Public|BindingFlags.Static)
                .Where(x => x.FieldType == typeof(Guid))
                .ToArray();

            foreach (var fieldInfo in allIdFields)
            {
                var id = (Guid?)fieldInfo.GetValue(null) ?? Guid.Empty;
                var found = BuiltInFieldId.Contains(id);
                Assert.AreEqual(true, found, $"Id {fieldInfo.Name} is defined on type BuiltInFieldId but not used in Contains.");
            }

        }
    }
}
