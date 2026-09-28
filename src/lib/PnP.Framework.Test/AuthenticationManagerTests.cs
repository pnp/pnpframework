using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Net;

namespace PnP.Framework.Test
{
    [TestClass]
    public class AuthenticationManagerTests
    {
        /// <summary>
        /// PnP Framework only forces a process wide TLS preference on the legacy targets. The
        /// guards that do so are compiled out from net9.0 onwards, and this test project runs on
        /// the newest supported target framework, so the preference has to be left alone here.
        /// A guard written for one single target framework silently starts applying again on the
        /// next one, which is what happened to net10.0 in issue #1218.
        /// </summary>
        [TestMethod]
        public void ConstructorDoesNotChangeProcessWideTlsPreference()
        {
#pragma warning disable SYSLIB0014 // ServicePointManager is obsolete, it is read here to assert it stays untouched
            SecurityProtocolType before = ServicePointManager.SecurityProtocol;

            using (new AuthenticationManager())
            {
            }

            Assert.AreEqual(before, ServicePointManager.SecurityProtocol);
#pragma warning restore SYSLIB0014
        }
    }
}
