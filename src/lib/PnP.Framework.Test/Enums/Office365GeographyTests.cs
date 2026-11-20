using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.Enums
{
    [TestClass]
    public class Office365GeographyTests
    {
        /// <summary>
        /// The PreferredDataLocation values from the table this enum documents itself against,
        /// https://learn.microsoft.com/microsoft-365/enterprise/multi-geo-add-group-with-pdl#geo-location-codes
        /// </summary>
        private static readonly string[] DocumentedGeoLocationCodes =
        {
            "APC", "AUS", "AUT", "BRA", "CAN", "CHL", "DNK", "EUR", "FRA", "DEU",
            "IND", "IDN", "ISR", "ITA", "JPN", "KOR", "MYS", "MEX", "NZL", "NOR",
            "POL", "QAT", "ZAF", "ESP", "SWE", "CHE", "TWN", "ARE", "GBR", "NAM",
        };

        /// <summary>
        /// Callers pass these codes straight through as PreferredDataLocation and the value is
        /// sent on with ToString(), so every documented code has to exist as a member and has to
        /// keep its own name.
        /// </summary>
        [TestMethod]
        public void EveryDocumentedGeoLocationCodeIsSupported()
        {
            List<string> missing = DocumentedGeoLocationCodes
                .Where(code => !Enum.TryParse(code, out Office365Geography parsed) || parsed.ToString() != code)
                .ToList();

            Assert.AreEqual(0, missing.Count, $"Not usable as PreferredDataLocation: {string.Join(", ", missing)}");
        }
    }
}
