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
    public class UsersUtilityTests
    {
        [TestMethod]
        public void ListUsers_WhenCalled_ShouldReturnData()
        {
            // Arrange
            TestCommon.RegisterPnPHttpClientMock();

            // Act
            var results = UsersUtility.ListUsers(
                "123",
                new[] {"postalCode"}
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
                    ),
                    pi.Name
                );
            }
        }
    }
}
