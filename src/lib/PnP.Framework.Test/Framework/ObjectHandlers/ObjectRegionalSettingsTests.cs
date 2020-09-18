using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectRegionalSettingsTests
    {
        Microsoft.SharePoint.Client.RegionalSettings defaultSettings;

        [TestInitialize]
        public void Initialize()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                defaultSettings = ctx.Web.RegionalSettings;
                ctx.Load(defaultSettings);
                ctx.Load(defaultSettings.TimeZone, tz => tz.Id);
                ctx.ExecuteQueryRetry();
            }
        }

        [TestMethod]
        public void CanExtractRegionalSettings()
        {
            using (var scope = new PnP.Framework.Diagnostics.PnPMonitoredScope("CanExtractRegionalSettings"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    // do not set the base template as that will mean that regional settings are not extracted
                    // when the test site has the same regional settings as the base template had
                    var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = null };

                    var template = new ProvisioningTemplate();
                    template = new ObjectRegionalSettings().ExtractObjects(ctx.Web, template, creationInfo);

                    Assert.IsNotNull(template.RegionalSettings);

                }
            }
        }

        [TestMethod]
        public void CanProvisionRegionalSettings()
        {
            using (var scope = new PnP.Framework.Diagnostics.PnPMonitoredScope("CanProvisionRegionalSettings"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var template = new ProvisioningTemplate
                    {
                        RegionalSettings = new PnP.Framework.Provisioning.Model.RegionalSettings
                        {
                            FirstDayOfWeek = System.DayOfWeek.Monday,
                            WorkDayEndHour = WorkHour.PM0700,
                            TimeZone = 5,
                            Time24 = true
                        }
                    };

                    var parser = new TokenParser(ctx.Web, template);
                    new ObjectRegionalSettings().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                    ctx.Load(ctx.Web.RegionalSettings);
                    ctx.Load(ctx.Web.RegionalSettings.TimeZone, tz => tz.Id);
                    ctx.ExecuteQueryRetry();

                    Assert.IsTrue(ctx.Web.RegionalSettings.Time24);
                    Assert.IsTrue(ctx.Web.RegionalSettings.WorkDayEndHour == (short)WorkHour.PM0700);
                    Assert.IsTrue(ctx.Web.RegionalSettings.FirstDayOfWeek == (uint)System.DayOfWeek.Monday);
                    Assert.IsTrue(ctx.Web.RegionalSettings.TimeZone.Id == 5);
                }
            }
        }

        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.RegionalSettings.AdjustHijriDays = defaultSettings.AdjustHijriDays;
                ctx.Web.RegionalSettings.AlternateCalendarType = defaultSettings.AlternateCalendarType;
                ctx.Web.RegionalSettings.CalendarType = defaultSettings.CalendarType;
                ctx.Web.RegionalSettings.Collation = defaultSettings.Collation;
                ctx.Web.RegionalSettings.FirstDayOfWeek = defaultSettings.FirstDayOfWeek;
                ctx.Web.RegionalSettings.FirstWeekOfYear = defaultSettings.FirstWeekOfYear;
                ctx.Web.RegionalSettings.LocaleId = defaultSettings.LocaleId;
                ctx.Web.RegionalSettings.ShowWeeks = defaultSettings.ShowWeeks;
                ctx.Web.RegionalSettings.Time24 = defaultSettings.Time24;
                ctx.Web.RegionalSettings.TimeZone = ctx.Web.RegionalSettings.TimeZones.GetById(defaultSettings.TimeZone.Id);
                ctx.Web.RegionalSettings.WorkDayEndHour = defaultSettings.WorkDayEndHour;
                ctx.Web.RegionalSettings.WorkDays = defaultSettings.WorkDays;
                ctx.Web.RegionalSettings.WorkDayStartHour = defaultSettings.WorkDayStartHour;
                ctx.Web.RegionalSettings.Update();
                ctx.Web.Update();
                ctx.ExecuteQueryRetry();
            }
        }
    }
}
