using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Entities;
using PnP.Framework.Graph;

namespace PnP.Framework.Test.Graph
{
    [TestClass]
    public class UnifiedGroupsUtilityTests
    {
        [TestMethod]
        public void GetUnifiedGroupMembers_WhenCalled_ShouldReturnData()
        {
            // Arrange
            TestCommon.RegisterPnPHttpClientMock();
            //var fakeHandler = new MockHttpHandler("");

            // Act
            var results = UnifiedGroupsUtility.GetUnifiedGroupMembers(
                new UnifiedGroupEntity() { GroupId = "abc", },
                "testt"
            );

            // Assert
            Assert.IsNotNull(results);
            Assert.IsTrue(results.Count > 1);

            AssertAllPropertiesHaveBeenAssigned(results);
        }

        private static void AssertAllPropertiesHaveBeenAssigned<T>(List<T> results)
        {
            foreach (PropertyInfo pi in typeof(T).GetProperties())
            {
                Assert.IsTrue(
                    results.Any(r =>
                        pi.GetValue(r, null) != (pi.PropertyType.IsValueType ? Activator.CreateInstance(pi.PropertyType) : null)
                    )
                );
            }
        }
    }
}
