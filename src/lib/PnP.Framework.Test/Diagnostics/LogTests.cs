using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Diagnostics;

namespace PnP.Framework.Tests.Diagnostics
{
    [TestClass]
    public class LogTests
    {
        [TestMethod]
        public void LogTest1()
        {
            Log.Info("Test Source", "Information test message");
            Log.Debug("Test Source", "Debug test message");

            Log.LogLevel = LogLevel.Information;

            Log.Error("Test Source", "Information test message 2");

        }
    }
}
