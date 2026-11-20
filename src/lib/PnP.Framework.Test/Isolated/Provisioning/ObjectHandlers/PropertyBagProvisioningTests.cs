using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Json;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace PnP.Framework.Test.Isolated.Provisioning.ObjectHandlers
{
    [TestClass]
    public class PropertyBagProvisioningTests
    {
        private const string TenantUrl = "https://propertybag-tests.sharepoint.invalid";
        private const string AdminUrl = "https://propertybag-tests-admin.sharepoint.invalid";

        [ClassInitialize]
        public static void Initialize(TestContext context)
        {
            // The hierarchy token parser reads the realm through WebRequest, outside CSOM.
            // Intercept only that endpoint on our fictional tenant, keeping the test offline.
            Assert.IsTrue(WebRequest.RegisterPrefix(new Uri(AdminUrl) + "/_vti_bin/client.svc", new RealmRequest()));
        }

        [DataTestMethod]
        [DataRow(true, false)]
        [DataRow(false, true)]
        public void ReusedOptionsDetectWritesSeparatelyForEachSiteCollection(bool firstAllowed, bool secondAllowed)
        {
            var options = new ProvisioningTemplateApplyingInformation
            {
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            };
            var firstResponses = new PropertyBagResponseProvider("/sites/first", firstAllowed);
            var secondResponses = new PropertyBagResponseProvider("/sites/second", secondAllowed);
            using var firstContext = CreateContext(firstResponses);
            using var secondContext = CreateContext(secondResponses);

            Assert.AreEqual(firstAllowed ? 2 : 1, InitializeTemplate(firstContext.Web, options));
            Assert.IsNull(options.PropertyBagWriteAllowed, "Detection must not become a caller override.");
            Assert.AreEqual(secondAllowed ? 2 : 1, InitializeTemplate(secondContext.Web, options));
            Assert.IsNull(options.PropertyBagWriteAllowed);
            Assert.AreEqual(1, firstResponses.ProbeWrites, "Handlers must share one probe result per site collection.");
            Assert.AreEqual(1, secondResponses.ProbeWrites);
            Assert.AreEqual(firstAllowed ? 1 : 0, firstResponses.ProbeCleanups);
            Assert.AreEqual(secondAllowed ? 1 : 0, secondResponses.ProbeCleanups);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void ExplicitOverrideDoesNotProbe(bool allowed)
        {
            var options = new ProvisioningTemplateApplyingInformation
            {
                PropertyBagWriteAllowed = allowed,
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            };
            var responses = new PropertyBagResponseProvider("/sites/override", !allowed);
            using var context = CreateContext(responses);

            Assert.AreEqual(allowed ? 2 : 1, InitializeTemplate(context.Web, options));
            Assert.AreEqual(allowed, options.PropertyBagWriteAllowed);
            Assert.AreEqual(0, responses.ProbeWrites, "An explicit override must skip the probe.");
            Assert.AreEqual(0, responses.ProbeCleanups);
        }

        [DataTestMethod]
        [DataRow(null, true, 1)]
        [DataRow(null, false, 0)]
        [DataRow(false, true, 0)]
        public void SkippedPropertyBagWritesWarnOnlyWhenWriteAccessCouldNotBeConfirmed(bool? callerOverride, bool hasWrites, int expectedWarnings)
        {
            var responses = new PropertyBagResponseProvider("/sites/blocked", false);
            using var context = CreateContext(responses);
            var template = hasWrites ? CreateTemplate() : new ProvisioningTemplate();
            template.Id = "Test {template}";
            var warnings = new List<string>();
            var completed = false;
            var options = new ProvisioningTemplateApplyingInformation
            {
                PropertyBagWriteAllowed = callerOverride,
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = hasWrites,
                MessagesDelegate = (message, type) =>
                {
                    if (type == ProvisioningMessageType.Warning)
                    {
                        warnings.Add(message);
                    }
                },
                SiteProvisionedDelegate = (title, url) => completed = true
            };

            new SiteToTemplateConversion().ApplyRemoteTemplate(context.Web, template, options);

            Assert.IsTrue(completed);
            Assert.AreEqual(0, responses.TemplateWrites);
            Assert.AreEqual(expectedWarnings, warnings.Count);
            if (expectedWarnings == 1)
            {
                var warning = warnings.Single();
                StringAssert.Contains(warning, responses.Url);
                StringAssert.Contains(warning, template.Id);
                StringAssert.Contains(warning, "property bag entries");
                StringAssert.Contains(warning, "template metadata");
            }
        }

        [TestMethod]
        public void ExplicitlyRequestedPropertyBagWriteStillPropagatesSharePointDenial()
        {
            var responses = new PropertyBagResponseProvider("/sites/denied", false);
            using var context = CreateContext(responses);
            var options = new ProvisioningTemplateApplyingInformation
            {
                PropertyBagWriteAllowed = true,
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            };

            Assert.ThrowsException<ServerUnauthorizedAccessException>(() =>
                new SiteToTemplateConversion().ApplyRemoteTemplate(context.Web, CreateTemplate(), options));
            Assert.AreEqual(0, responses.ProbeWrites);
            Assert.AreEqual(1, responses.TemplateWrites);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void PropertyBagProbeDoesNotExecutePendingCallerUpdates(bool loadPermissions)
        {
            var responses = new PropertyBagResponseProvider("/sites/queued", true) { FailCallerUpdate = true };
            using var context = CreateContext(responses);
            if (loadPermissions)
            {
                Assert.IsTrue(context.Web.IsNoScriptSite());
            }
            context.Web.Title = "Caller update";
            context.Web.Update();

            Assert.IsFalse(context.Web.IsPropertyBagWriteBlocked());

            Assert.AreEqual(1, responses.ProbeWrites);
            Assert.AreEqual(0, responses.CallerUpdates);
            Assert.IsTrue(context.HasPendingRequest);
            Assert.ThrowsException<ServerUnauthorizedAccessException>(() => context.ExecuteQueryRetry());
            Assert.AreEqual(1, responses.CallerUpdates);
        }

        [TestMethod]
        public void PropertyBagProbeAndTemplateWriteUseTheCallersCredentials()
        {
            var responses = new PropertyBagResponseProvider("/sites/credentials", true);
            var credentials = new NetworkCredential("probe-user", "test");
            using var context = new ClientContext(responses.Url)
            {
                Credentials = credentials,
                WebRequestExecutorFactory = new CredentialCheckingRequestExecutorFactory(responses, credentials)
            };
            context.Web.ApplyProvisioningTemplate(CreateTemplate(), new ProvisioningTemplateApplyingInformation
            {
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            });

            Assert.AreEqual(1, responses.ProbeWrites);
            Assert.AreEqual(1, responses.ProbeCleanups);
            Assert.AreEqual(1, responses.TemplateWrites);
        }

        [TestMethod]
        public void FailedProbeCleanupStillReportsAConfirmedWrite()
        {
            var responses = new PropertyBagResponseProvider("/sites/cleanup", true) { FailProbeCleanup = true };
            using var context = CreateContext(responses);

            Assert.IsFalse(context.Web.IsPropertyBagWriteBlocked());

            Assert.AreEqual(1, responses.ProbeWrites);
            Assert.AreEqual(1, responses.ProbeCleanups);
        }

        [DataTestMethod]
        [DataRow("RootWeb")]
        [DataRow("OpenWeb")]
        [DataRow("OpenWebById")]
        [DataRow("ListByTitle")]
        [DataRow("RootListById")]
        public void StandaloneProbeResolvesAnUnloadedWebWithoutRunningCallerUpdates(string lookup)
        {
            bool targetsRoot = lookup == "RootWeb" || lookup == "RootListById";
            var responses = new PropertyBagResponseProvider("/sites/probe", !targetsRoot, webPath: "/sites/probe/subsite", rootWritesAllowed: targetsRoot)
            {
                FailCallerUpdate = true
            };
            using var context = CreateContext(responses);
            context.DisableReturnValueCache = true;
            var web = lookup switch
            {
                "RootWeb" => context.Site.RootWeb,
                "OpenWeb" => context.Site.OpenWeb("subsite"),
                "ListByTitle" => context.Web.Lists.GetByTitle("Documents").ParentWeb,
                "RootListById" => context.Site.RootWeb.Lists.GetById(Guid.Parse("22222222-2222-2222-2222-222222222222")).ParentWeb,
                _ => context.Site.OpenWebById(Guid.Parse("11111111-1111-1111-1111-111111111111"))
            };
            Assert.IsTrue(context.HasPendingRequest);
            Assert.IsFalse(web.IsPropertyAvailable(w => w.Url));
            context.Web.Title = "Caller update";
            context.Web.Update();

            Assert.IsFalse(web.IsPropertyBagWriteBlocked());

            Assert.AreEqual(1, responses.ProbeWrites);
            Assert.AreEqual(1, responses.ProbeCleanups);
            Assert.AreEqual(0, responses.CallerUpdates);
            Assert.IsTrue(context.HasPendingRequest);
            Assert.ThrowsException<ServerUnauthorizedAccessException>(() => context.ExecuteQueryRetry());
        }

        [DataTestMethod]
        [DataRow("2019/03", false)]
        [DataRow("2019/03", true)]
        [DataRow("2022/09", false)]
        [DataRow("2022/09", true)]
        public void XmlRoundTripDistinguishesOmittedNoCrawlFromFalse(string schema, bool specified)
        {
            string ns = $"http://schemas.dev.office.com/PnP/{schema}/ProvisioningSchema";
            var formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(ns);
            string noCrawl = specified ? " NoCrawl=\"false\"" : string.Empty;
            string xml = $"<ProvisioningTemplate xmlns=\"{ns}\" ID=\"Test\"><WebSettings Title=\"Renamed site\"{noCrawl} /></ProvisioningTemplate>";
            using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            var template = formatter.ToProvisioningTemplate(source);

            Assert.AreEqual(specified, template.WebSettings.NoCrawlSpecified);
            Assert.IsFalse(template.WebSettings.NoCrawl);
            var jsonFormatter = new JsonPnPFormatter();
            using var json = jsonFormatter.ToFormattedTemplate(template);
            var fromJson = jsonFormatter.ToProvisioningTemplate(json);
            Assert.AreEqual(specified, fromJson.WebSettings.NoCrawlSpecified);
            using var output = formatter.ToFormattedTemplate(fromJson);
            var settings = XDocument.Load(output).Descendants(XName.Get("WebSettings", ns)).Single();
            Assert.AreEqual(specified ? "false" : null, (string)settings.Attribute("NoCrawl"));
        }

        [DataTestMethod]
        [DataRow(false, false)]
        [DataRow(false, true)]
        [DataRow(true, true)]
        public void ApplyWebSettingsOnlyChangesNoCrawlWhenSpecifiedOnNoScript(bool explicitNoCrawl, bool useJson)
        {
            var template = new ProvisioningTemplate { WebSettings = new WebSettings { Title = "Renamed site" } };
            if (explicitNoCrawl)
            {
                template.WebSettings.NoCrawl = false;
            }
            if (useJson)
            {
                var formatter = new JsonPnPFormatter();
                using var json = formatter.ToFormattedTemplate(template);
                template = formatter.ToProvisioningTemplate(json);
            }

            var responses = new PropertyBagResponseProvider("/sites/crawl", true);
            using var context = CreateContext(responses);
            new ObjectWebSettings().ProvisionObjects(context.Web, template, new TokenParser(context.Web, template),
                new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = true });

            Assert.AreEqual(!explicitNoCrawl, context.Web.NoCrawl);
            Assert.AreEqual(explicitNoCrawl ? 1 : 0, responses.NoCrawlUpdates);
            Assert.AreEqual("Renamed site", context.Web.Title);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void StandaloneSubwebChecksItsOwnWriteAccessBeforeCountingHandlers(bool allowed)
        {
            var responses = new PropertyBagResponseProvider("/sites/progress", allowed, "/sites/progress/subsite", rootWritesAllowed: !allowed, canCustomizePages: allowed);
            using var context = CreateContext(responses);
            var options = new ProvisioningTemplateApplyingInformation
            {
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            };

            Assert.AreEqual(allowed ? 2 : 1, InitializeTemplate(context.Web, options), "Progress must count only handlers that can run.");
            Assert.AreEqual(allowed ? 0 : 1, responses.ProbeWrites, "A subweb with AddAndCustomizePages permission needs no probe.");
            Assert.AreEqual(0, responses.ProbeCleanups);
            Assert.AreEqual(0, responses.RootProbeRequests, "A subweb's permissions must not be inferred from its root web.");
            Assert.IsNull(options.PropertyBagWriteAllowed);
        }

        [DataTestMethod]
        [DataRow(true, null)]
        [DataRow(false, null)]
        [DataRow(true, true)]
        [DataRow(true, false)]
        public void HierarchyHonorsOverridesOrChecksEachWebOnce(bool allowed, bool? callerOverride)
        {
            // Auto-detection must handle different permissions within and across collections.
            // Explicit overrides are exercised on writable webs to verify both opting in and opting out.
            bool otherAllowed = callerOverride.HasValue ? allowed : !allowed;
            var firstResponses = new PropertyBagResponseProvider("/sites/first", allowed, includeSubweb: true);
            var firstSubwebResponses = new PropertyBagResponseProvider("/sites/first", otherAllowed, "/sites/first/subsite");
            var secondResponses = new PropertyBagResponseProvider("/sites/second", otherAllowed, includeSubweb: true);
            var secondSubwebResponses = new PropertyBagResponseProvider("/sites/second", allowed, "/sites/second/subsite");
            firstResponses.Subweb = firstSubwebResponses;
            secondResponses.Subweb = secondSubwebResponses;
            using var context = new ClientContext(AdminUrl)
            {
                WebRequestExecutorFactory = new MockWebRequestExecutorFactory(new HierarchyResponseProvider(
                    firstResponses, firstSubwebResponses, secondResponses, secondSubwebResponses))
            };

            var hierarchy = new ProvisioningHierarchy();
            var template = CreateTemplate();
            template.Id = "PropertyBag";
            hierarchy.Templates.Add(template);
            var sequence = new ProvisioningSequence { ID = "Sites" };
            hierarchy.Sequences.Add(sequence);
            foreach (var responses in new[] { firstResponses, secondResponses })
            {
                sequence.SiteCollections.Add(new ClassicSiteCollection
                {
                    Url = responses.Url,
                    Title = responses.Url,
                    Templates = { template.Id, template.Id },
                    Sites = { new TeamNoGroupSubSite { Url = "subsite", Templates = { template.Id, template.Id } } }
                });
            }

            var completedUrls = new List<string>();
            var warnings = new List<string>();
            var configuration = new ApplyConfiguration();
            configuration.Handlers.Add(ConfigurationHandler.PropertyBagEntries);
            configuration.PropertyBag.PropertyBagWriteAllowed = callerOverride;
            configuration.SiteProvisionedDelegate = (title, url) => completedUrls.Add(url);
            configuration.MessagesDelegate = (message, type) =>
            {
                if (type == ProvisioningMessageType.Warning)
                {
                    warnings.Add(message);
                }
            };
            var tenant = new Tenant(context);
            var parser = new TokenParser(tenant, hierarchy);

            // Exercise the real loop: the test must not create the per-collection applying options itself.
            new ObjectHierarchySequenceSites().ProvisionObjects(tenant, hierarchy, sequence.ID, parser, configuration);

            CollectionAssert.AreEqual(new[]
            {
                firstResponses.Url, firstResponses.Url, firstSubwebResponses.Url, firstSubwebResponses.Url,
                secondResponses.Url, secondResponses.Url, secondSubwebResponses.Url, secondSubwebResponses.Url
            }, completedUrls);
            var expectedProbes = callerOverride.HasValue ? 0 : 1;
            Assert.AreEqual(expectedProbes, firstResponses.ProbeWrites);
            Assert.AreEqual(expectedProbes, secondResponses.ProbeWrites);
            Assert.AreEqual(expectedProbes, firstSubwebResponses.ProbeWrites, "A subweb must check its own permissions unless the caller overrides detection.");
            Assert.AreEqual(expectedProbes, secondSubwebResponses.ProbeWrites);
            Assert.AreEqual((callerOverride ?? allowed) ? 2 : 0, firstResponses.TemplateWrites);
            Assert.AreEqual((callerOverride ?? otherAllowed) ? 2 : 0, firstSubwebResponses.TemplateWrites);
            Assert.AreEqual((callerOverride ?? otherAllowed) ? 2 : 0, secondResponses.TemplateWrites);
            Assert.AreEqual((callerOverride ?? allowed) ? 2 : 0, secondSubwebResponses.TemplateWrites);

            // Both applications on each blocked web should warn, even though they share one probe.
            Assert.AreEqual(callerOverride.HasValue ? 0 : 4, warnings.Count);
            if (!callerOverride.HasValue)
            {
                var blockedRootUrl = allowed ? secondResponses.Url : firstResponses.Url;
                var blockedSubwebUrl = allowed ? firstSubwebResponses.Url : secondSubwebResponses.Url;
                Assert.AreEqual(2, warnings.Count(message => message.Contains($"'{blockedRootUrl}'")));
                Assert.AreEqual(2, warnings.Count(message => message.Contains($"'{blockedSubwebUrl}'")));
            }
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow(true)]
        [DataRow(false)]
        public void HierarchyConfigurationPreservesCallerOverride(bool? allowed)
        {
            var options = new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = allowed };
            var configuration = ApplyConfiguration.FromApplyingInformation(options);

            Assert.AreEqual(allowed, configuration.ToApplyingInformation().PropertyBagWriteAllowed);
            var json = JsonSerializer.Serialize(new { propertyBag = new { propertyBagWriteAllowed = allowed } });
            var fromJson = ApplyConfiguration.FromString(json);
            Assert.AreEqual(allowed, fromJson.ToApplyingInformation().PropertyBagWriteAllowed);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void ComposedLookHonorsPropertyBagOverrideWhileApplyingTheme(bool allowed)
        {
            var responses = new PropertyBagResponseProvider("/sites/branding", true, canCustomizePages: true);
            responses.AddComposedLookResponses();
            using var context = CreateContext(responses);
            var template = new ProvisioningTemplate
            {
                ComposedLook = new ComposedLook { Name = "Test look", ColorFile = "/_catalogs/theme/15/test.spcolor" }
            };
            var parser = new TokenParser(context.Web, template);

            new ObjectComposedLook().ProvisionObjects(context.Web, template, parser,
                new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = allowed });

            Assert.AreEqual(1, responses.ThemeApplications, "Disabling property bag writes must still apply the theme.");
            Assert.AreEqual(2, responses.MasterPageUpdates, "Both master URLs must still be applied.");
            Assert.AreEqual(allowed ? 2 : 0, responses.MasterPageInheritanceWrites);
            Assert.AreEqual(allowed ? 1 : 0, responses.ThemeInheritanceWrites);
            Assert.AreEqual(allowed ? 1 : 0, responses.ComposedLookMetadataWrites);
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("palette")]
        [DataRow("font")]
        public void ExtractionUsesTheGalleryWhenComposedLookMetadataIsStale(string changedAsset)
        {
            var responses = new PropertyBagResponseProvider("/sites/branding", true, canCustomizePages: true);
            responses.AddComposedLookResponses();
            var handler = new ObjectComposedLook();
            var first = new ProvisioningTemplate
            {
                ComposedLook = new ComposedLook { Name = "Test look", ColorFile = "{parameter:Palette}", Version = 7 }
            };
            first.Parameters.Add("Palette", "/sites/branding/SiteAssets/first.spcolor");
            first.Parameters.Add("Unrelated", "unused");
            using (var context = CreateContext(responses))
            {
                handler.ProvisionObjects(context.Web, first, new TokenParser(context.Web, first),
                    new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = true });
            }

            if (changedAsset != null)
            {
                // Change the palette but keep the name. The stored metadata stays on the first version.
                var second = new ProvisioningTemplate
                {
                    ComposedLook = new ComposedLook
                    {
                        Name = "Test look",
                        ColorFile = changedAsset == "palette" ? "/sites/branding/SiteAssets/second.spcolor" : first.Parameters["Palette"],
                        FontFile = changedAsset == "font" ? "/sites/branding/SiteAssets/added.spfont" : null
                    }
                };
                using var context = CreateContext(responses);
                handler.ProvisionObjects(context.Web, second, new TokenParser(context.Web, second),
                    new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = false });
            }

            using var extractionContext = CreateContext(responses);
            var extracted = handler.ExtractObjects(extractionContext.Web, new ProvisioningTemplate(), null);

            Assert.AreEqual(1, responses.ComposedLookMetadataWrites);
            Assert.AreEqual(changedAsset == null ? 7 : 0, extracted.ComposedLook.Version);
            if (changedAsset == null)
            {
                Assert.AreEqual("{parameter:Palette}", extracted.ComposedLook.ColorFile);
                Assert.AreEqual(first.Parameters["Palette"], extracted.Parameters["Palette"]);
                Assert.IsFalse(extracted.Parameters.ContainsKey("Unrelated"));
            }
            else
            {
                Assert.AreEqual(changedAsset == "palette" ? "{site}/SiteAssets/second.spcolor" : "{site}/SiteAssets/first.spcolor",
                    extracted.ComposedLook.ColorFile);
            }
            var expectedFiles = changedAsset == "font" ? new[] { "first.spcolor", "added.spfont" }
                : new[] { changedAsset == "palette" ? "second.spcolor" : "first.spcolor" };
            CollectionAssert.AreEqual(expectedFiles,
                extracted.Files.Select(file => file.Src).ToArray());
            Assert.IsTrue(extracted.Files.All(file => file.Folder == "{site}/SiteAssets"));
        }

        [DataTestMethod]
        [DataRow(true, false)]
        [DataRow(true, true)]
        [DataRow(false, true)]
        public void ExtractionPreservesMatchingMetadataWhenTheGalleryContainsAliases(bool legacyMetadata, bool alias)
        {
            var responses = new PropertyBagResponseProvider("/sites/branding", true, canCustomizePages: true) { DuplicateLook = alias };
            responses.AddComposedLookResponses();
            var handler = new ObjectComposedLook();
            var template = new ProvisioningTemplate
            {
                ComposedLook = new ComposedLook { Name = "Test look", ColorFile = "{site}/SiteAssets/first.spcolor", Version = 7 }
            };
            using var context = CreateContext(responses);
            handler.ProvisionObjects(context.Web, template, new TokenParser(context.Web, template),
                new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = true });
            if (legacyMetadata)
            {
                responses.SetComposedLookMetadata("{\"Name\":\"Test look\",\"ColorFile\":\"{site}/SiteAssets/first.spcolor\",\"Version\":7}");
            }

            using var extractionContext = CreateContext(responses);
            var exported = handler.ExtractObjects(extractionContext.Web, new ProvisioningTemplate(), null);

            Assert.AreEqual("Test look", exported.ComposedLook.Name);
            Assert.AreEqual(7, exported.ComposedLook.Version);
            Assert.AreEqual("{site}/SiteAssets/first.spcolor", exported.ComposedLook.ColorFile);
        }

        [TestMethod]
        public void ExtractionWithoutStoredMetadataDistinguishesLooksByBackground()
        {
            var responses = new PropertyBagResponseProvider("/sites/branding", true, canCustomizePages: true)
            {
                OtherLookBackground = "/sites/branding/SiteAssets/other.png"
            };
            responses.AddComposedLookResponses();
            using (var context = CreateContext(responses))
            {
                var template = new ProvisioningTemplate
                {
                    ComposedLook = new ComposedLook
                    {
                        Name = "Test look",
                        ColorFile = "/sites/branding/SiteAssets/first.spcolor",
                        BackgroundFile = "/sites/branding/SiteAssets/active.png"
                    }
                };
                new ObjectComposedLook().ProvisionObjects(context.Web, template, new TokenParser(context.Web, template),
                    new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = false });
            }

            using var extractionContext = CreateContext(responses);
            var exported = new ObjectComposedLook().ExtractObjects(extractionContext.Web, new ProvisioningTemplate(), null);
            Assert.AreEqual("Test look", exported.ComposedLook.Name);
            Assert.AreEqual("{site}/SiteAssets/active.png", exported.ComposedLook.BackgroundFile);
        }

        [TestMethod]
        public void RecursiveThemeChecksTheChildBeforeWritingItsInheritanceProperty()
        {
            var root = new PropertyBagResponseProvider("/sites/theme", true, includeSubweb: true, canCustomizePages: true);
            root.Subweb = new PropertyBagResponseProvider("/sites/theme", false, "/sites/theme/subsite");
            root.AddComposedLookResponses();
            root.Subweb.AddComposedLookResponses();
            using var context = CreateContext(root);

            context.Web.SetThemeByUrl("/_catalogs/theme/15/test.spcolor", null, null, resetSubsitesToInherit: true);

            Assert.AreEqual(1, root.ThemeInheritanceWrites);
            Assert.AreEqual(1, root.Subweb.ProbeWrites);
            Assert.AreEqual(0, root.Subweb.ThemeInheritanceWrites);
            Assert.AreEqual(1, root.Subweb.ThemeApplications, "A blocked inheritance write must not skip the theme.");
        }

        [TestMethod]
        public void PublishingSkipsPropertyBagSettingsWhenWritesAreDisabled()
        {
            var responses = new PropertyBagResponseProvider("/sites/publishing", true, canCustomizePages: true);
            responses.AddPublishingFeatures();
            using var context = CreateContext(responses);
            var template = new ProvisioningTemplate { Publishing = new Publishing() };
            template.Publishing.AvailableWebTemplates.Add(new AvailableWebTemplate { TemplateName = "STS#0", LanguageCode = 1033 });
            template.Publishing.PageLayouts.Add(new PageLayout { Path = "ArticleLeft.aspx", IsDefault = true });
            var parser = new TokenParser(context.Web, template);

            var skippedSettings = new List<string>();
            var handler = new ObjectPublishing { ReportSkippedPropertyBagWrite = skippedSettings.Add };
            handler.ProvisionObjects(context.Web, template, parser,
                new ProvisioningTemplateApplyingInformation { PropertyBagWriteAllowed = false });

            Assert.AreEqual(2, responses.PublishingFeatureActivations);
            Assert.AreEqual(0, responses.PublishingWrites);
            CollectionAssert.AreEqual(new[] { "available web templates", "publishing page layout settings" }, skippedSettings);
        }

        private static int? InitializeTemplate(Web web, ProvisioningTemplateApplyingInformation options)
        {
            int? handlerCount = null;
            options.ProgressDelegate = (message, step, total) =>
            {
                handlerCount = total;
                // Stop after the real pipeline's first WillProvision calls, before applying the template.
                throw new StopAfterInitializationException();
            };
            Assert.ThrowsException<StopAfterInitializationException>(() =>
                new SiteToTemplateConversion().ApplyRemoteTemplate(web, CreateTemplate(), options));
            return handlerCount;
        }

        private static ProvisioningTemplate CreateTemplate()
        {
            var template = new ProvisioningTemplate();
            template.PropertyBagEntries.Add(new PropertyBagEntry { Key = "Example", Value = "Value", Overwrite = true });
            return template;
        }

        private static ClientContext CreateContext(PropertyBagResponseProvider responses)
        {
            return new ClientContext(responses.Url)
            {
                WebRequestExecutorFactory = new MockWebRequestExecutorFactory(responses)
            };
        }

        private sealed class StopAfterInitializationException : Exception { }

        private sealed class CredentialCheckingRequestExecutorFactory : MockWebRequestExecutorFactory
        {
            private readonly NetworkCredential expected;

            public CredentialCheckingRequestExecutorFactory(IMockResponseProvider responses, NetworkCredential expected)
                : base(responses)
            {
                this.expected = expected;
            }

            public override WebRequestExecutor CreateWebRequestExecutor(ClientRuntimeContext context, string requestUrl)
            {
                var credential = context.Credentials?.GetCredential(new Uri(requestUrl), "Basic");
                if (credential?.UserName != expected.UserName || credential.Password != expected.Password)
                {
                    throw new UnauthorizedAccessException("The request context must supply the test user's credentials.");
                }
                return base.CreateWebRequestExecutor(context, requestUrl);
            }
        }

#pragma warning disable SYSLIB0014 // Match the legacy WebRequest API used by GetAuthenticationRealm.
        private sealed class RealmRequest : WebRequest, IWebRequestCreate
        {
            public override WebHeaderCollection Headers { get; set; } = new WebHeaderCollection();
            WebRequest IWebRequestCreate.Create(Uri uri) => new RealmRequest();
            public override WebResponse GetResponse() => new RealmResponse();
        }
#pragma warning restore SYSLIB0014

        private sealed class RealmResponse : WebResponse { }

        private sealed class HierarchyResponseProvider : IMockResponseProvider
        {
            private readonly PropertyBagResponseProvider[] sites;
            private readonly MockEntryResponseProvider admin = new MockEntryResponseProvider();

            public HierarchyResponseProvider(params PropertyBagResponseProvider[] sites)
            {
                this.sites = sites;
                admin.ResponseEntries.Add(new MockResponseEntry { Url = AdminUrl, PropertyName = "Site", ReturnValue = new { Url = AdminUrl } });
                admin.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = AdminUrl,
                    PropertyName = "Web",
                    ReturnValue = new
                    {
                        Url = AdminUrl,
                        ServerRelativeUrl = "/",
                        Language = 1033,
                        CurrentUser = new { _ObjectType_ = "SP.User", Id = 1, LoginName = "testuser", Title = "Test User" }
                    }
                });
                admin.ResponseEntries.Add(new MockResponseEntry { Url = AdminUrl, Method = "GetAllTenantThemes", ReturnValue = new { _Child_Items_ = Array.Empty<object>() } });
                admin.ResponseEntries.Add(new MockResponseEntry { Url = AdminUrl, Method = "GetSitePropertiesByUrl", ReturnValue = new { Status = "Active" } });
            }

            public string GetResponse(string url, string verb, string body)
            {
                if (url == AdminUrl + "/_vti_bin/client.svc/ProcessQuery")
                {
                    return admin.GetResponse(url, verb, body);
                }
                var site = sites.Single(s => url == s.Url + "/_vti_bin/client.svc/ProcessQuery");
                return site.GetResponse(url, verb, body);
            }
        }

        private sealed class PropertyBagResponseProvider : IMockResponseProvider
        {
            private readonly MockEntryResponseProvider reads = new MockEntryResponseProvider();
            private readonly bool writesAllowed;
            private readonly bool rootWritesAllowed;
            private const string WebId = "/Guid(11111111-1111-1111-1111-111111111111)/";
            private const string ComposedLookInfoKey = "_PnP_ProvisioningTemplateComposedLookInfo";
            private MockResponseEntry gallery;
            private string palette = "/_catalogs/theme/15/test.spcolor";
            private string font;
            private string background;
            private const string AccessDeniedResponse = "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\"," +
                "\"ErrorInfo\":{\"ErrorMessage\":\"Access denied\",\"ErrorCode\":-2147024891," +
                "\"ErrorTypeName\":\"System.UnauthorizedAccessException\"}}]";

            public string Url { get; }
            public int ProbeWrites { get; private set; }
            public int ProbeCleanups { get; private set; }
            public int RootProbeRequests { get; private set; }
            public int TemplateWrites { get; private set; }
            public PropertyBagResponseProvider Subweb { get; set; }
            public int ThemeApplications { get; private set; }
            public int ThemeInheritanceWrites { get; private set; }
            public int ComposedLookMetadataWrites { get; private set; }
            public int MasterPageUpdates { get; private set; }
            public int MasterPageInheritanceWrites { get; private set; }
            public int PublishingWrites { get; private set; }
            public int PublishingFeatureActivations { get; private set; }
            public bool FailCallerUpdate { get; set; }
            public bool FailProbeCleanup { get; set; }
            public int CallerUpdates { get; private set; }
            public int NoCrawlUpdates { get; private set; }
            public bool DuplicateLook { get; set; }
            public string OtherLookBackground { get; set; }

            public PropertyBagResponseProvider(string path, bool writesAllowed, string webPath = null, bool includeSubweb = false, bool? rootWritesAllowed = null, bool canCustomizePages = false)
            {
                Url = TenantUrl + (webPath ?? path);
                this.writesAllowed = writesAllowed;
                this.rootWritesAllowed = rootWritesAllowed ?? writesAllowed;
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "Web",
                    ReturnValue = CreateWebResponse(webPath ?? path, includeSubweb, canCustomizePages: canCustomizePages)
                });
                if (includeSubweb)
                {
                    reads.ResponseEntries.Add(new MockResponseEntry { Url = Url, PropertyName = "SubWeb", ReturnValue = CreateWebResponse(path + "/subsite", false, true) });
                }
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "RootWeb",
                    ReturnValue = new
                    {
                        Id = WebId,
                        Url = TenantUrl + path,
                        EffectiveBasePermissions = new { _ObjectType_ = "SP.BasePermissions", High = 0, Low = 0 }
                    }
                });
                foreach (var method in new[] { "OpenWeb", "OpenWebById" })
                {
                    reads.ResponseEntries.Add(new MockResponseEntry
                    {
                        Url = Url, Method = method,
                        ReturnValue = CreateWebResponse(webPath ?? path, includeSubweb, canCustomizePages: canCustomizePages)
                    });
                }
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "AllProperties",
                    ReturnValue = new { _ObjectType_ = "SP.PropertyValues", ExistingKey = "ExistingValue" }
                });
            }

            public void AddComposedLookResponses()
            {
                gallery = new MockResponseEntry { Url = Url, Method = "GetItems" };
                reads.ResponseEntries.Add(gallery);
                UpdateGalleryResponse();
                reads.ResponseEntries.Add(new MockResponseEntry { Url = Url, PropertyName = "Site", ReturnValue = new { Url } });
                reads.ResponseEntries.Add(new MockResponseEntry { Url = Url, Method = "ApplyTheme", ReturnValue = null });
            }

            public void SetComposedLookMetadata(string metadata)
            {
                reads.ResponseEntries.Single(entry => entry.PropertyName == "AllProperties").ReturnValue =
                    new Dictionary<string, object> { ["_ObjectType_"] = "SP.PropertyValues", [ComposedLookInfoKey] = metadata };
            }

            public void AddPublishingFeatures()
            {
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "Features",
                    ReturnValue = new
                    {
                        _Child_Items_ = Array.Empty<object>()
                    }
                });
                reads.ResponseEntries.Add(new MockResponseEntry { Url = Url, Method = "GetById", ReturnValue = new { IsNull = true } });
            }

            private void UpdateGalleryResponse()
            {
                var items = new List<object>();
                if (OtherLookBackground != null || DuplicateLook)
                {
                    items.Add(GalleryItem("Other look", 3, DuplicateLook ? background : OtherLookBackground));
                }
                items.Add(GalleryItem("Test look", 1, background));
                items.Add(GalleryItem("Current", 2, background));
                gallery.ReturnValue = new { _Child_Items_ = items };
            }

            private object GalleryItem(string name, int id, string image)
            {
                return new
                {
                    _ObjectType_ = "SP.ListItem", _ObjectIdentity_ = $"look{id}", Id = id, Name = name,
                    ThemeUrl = new { _ObjectType_ = "SP.FieldUrlValue", Url = TenantUrl + palette, Description = "Palette" },
                    FontSchemeUrl = string.IsNullOrEmpty(font) ? null : new { _ObjectType_ = "SP.FieldUrlValue", Url = TenantUrl + font },
                    ImageUrl = string.IsNullOrEmpty(image) ? null : new { _ObjectType_ = "SP.FieldUrlValue", Url = TenantUrl + image },
                    MasterPageUrl = new { _ObjectType_ = "SP.FieldUrlValue", Url = Url + "/_catalogs/masterpage/seattle.master", Description = "Master page" }
                };
            }

            private static object CreateWebResponse(string path, bool includeSubweb, bool isSubweb = false, bool canCustomizePages = false)
            {
                return new
                {
                    _ObjectType_ = "SP.Web",
                    _ObjectIdentity_ = isSubweb ? "SubWeb" : "Web",
                    Id = WebId,
                    Url = TenantUrl + path,
                    ServerRelativeUrl = path,
                    Title = path,
                    Language = 1033,
                    WebTemplate = "STS",
                    Configuration = 0,
                    NoCrawl = true,
                    CommentsOnSitePagesDisabled = false,
                    ExcludeFromOfflineClient = false,
                    MembersCanShare = false,
                    DisableFlows = false,
                    DisableAppViews = false,
                    HorizontalQuickLaunch = false,
                    SearchScope = 0,
                    SearchBoxInNavBar = 0,
                    EffectiveBasePermissions = new { _ObjectType_ = "SP.BasePermissions", High = 0,
                        Low = canCustomizePages ? 1L << ((int)PermissionKind.AddAndCustomizePages - 1) : 0L },
                    Webs = new { _Child_Items_ = includeSubweb ? new[] { CreateWebResponse(path + "/subsite", false, true) } : Array.Empty<object>() }
                };
            }

            public string GetResponse(string url, string verb, string body)
            {
                if (body.Contains("Name=\"GetLocalizedString\""))
                {
                    // The shared mock does not return values for static method actions.
                    var action = XDocument.Parse(body).Descendants().Single(element => (string)element.Attribute("Name") == "GetLocalizedString");
                    return $"[{{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\",\"ErrorInfo\":null}},{(int)action.Attribute("Id")},\"Current\"]";
                }
                if (Subweb != null && (url == Subweb.Url + "/_vti_bin/client.svc/ProcessQuery" ||
                    (body.Contains("Name=\"SubWeb\"") && !body.Contains("Name=\"RootWeb\""))))
                {
                    return Subweb.GetResponse(Subweb.Url + "/_vti_bin/client.svc/ProcessQuery", verb, body);
                }
                if (body.Contains("<SetProperty") && body.Contains("Name=\"Title\""))
                {
                    if (FailCallerUpdate)
                    {
                        CallerUpdates++;
                        return AccessDeniedResponse;
                    }
                    if (body.Contains("Name=\"NoCrawl\""))
                    {
                        NoCrawlUpdates++;
                    }
                    return "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\",\"ErrorInfo\":null}]";
                }
                if (body.Contains("Name=\"Features\"") && body.Contains("Name=\"Add\""))
                {
                    PublishingFeatureActivations++;
                }
                if (body.Contains("<SetProperty") && (body.Contains("Name=\"MasterUrl\"") || body.Contains("Name=\"CustomMasterUrl\"")))
                {
                    MasterPageUpdates++;
                    if (!body.Contains("Name=\"SetFieldValue\""))
                    {
                        return "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\",\"ErrorInfo\":null}]";
                    }
                }
                if (body.Contains("Name=\"ApplyTheme\""))
                {
                    ThemeApplications++;
                    var assets = MethodParameters(body, "ApplyTheme").Take(3).Select(parameter => parameter.Value).ToArray();
                    palette = assets[0];
                    font = assets[1];
                    background = assets[2];
                    UpdateGalleryResponse();
                }
                if (body.Contains("Name=\"SetFieldValue\""))
                {
                    if (body.Contains("_PnP_PropertyBagProbe_"))
                    {
                        if (body.Contains("Name=\"RootWeb\""))
                        {
                            RootProbeRequests++;
                        }
                        if (body.Contains("Type=\"Null\""))
                        {
                            ProbeCleanups++;
                            if (FailProbeCleanup)
                            {
                                return AccessDeniedResponse;
                            }
                        }
                        else
                        {
                            ProbeWrites++;
                        }
                    }
                    else if (body.Contains(">Example</Parameter>"))
                    {
                        TemplateWrites++;
                    }
                    else if (body.Contains(">__InheritsThemedCssFolderUrl</Parameter>"))
                    {
                        ThemeInheritanceWrites++;
                    }
                    else if (body.Contains(">_PnP_ProvisioningTemplateComposedLookInfo</Parameter>"))
                    {
                        ComposedLookMetadataWrites++;
                        SetComposedLookMetadata(MethodParameters(body, "SetFieldValue").Last().Value);
                    }
                    else if (body.Contains(">__InheritsMasterUrl</Parameter>") || body.Contains(">__InheritsCustomMasterUrl</Parameter>"))
                    {
                        MasterPageInheritanceWrites++;
                    }
                    else if (body.Contains(">__PageLayouts</Parameter>") || body.Contains(">__DefaultPageLayout</Parameter>") ||
                        body.Contains(">__WebTemplates</Parameter>") || body.Contains(">__InheritWebTemplates</Parameter>"))
                    {
                        PublishingWrites++;
                    }

                    if (!(body.Contains("Name=\"RootWeb\"") ? rootWritesAllowed : writesAllowed))
                    {
                        return AccessDeniedResponse;
                    }

                    return "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\",\"ErrorInfo\":null}]";
                }

                return reads.GetResponse(url, verb, body);
            }

            private static IEnumerable<XElement> MethodParameters(string body, string name)
            {
                return XDocument.Parse(body).Descendants()
                    .First(element => element.Name.LocalName == "Method" && (string)element.Attribute("Name") == name)
                    .Elements().Single(element => element.Name.LocalName == "Parameters").Elements();
            }
        }
    }
}
