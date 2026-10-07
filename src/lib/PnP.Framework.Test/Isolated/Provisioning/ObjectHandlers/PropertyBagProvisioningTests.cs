using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

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
        [DataRow(true)]
        [DataRow(false)]
        public void StandaloneSubwebProbesRootWebBeforeCountingHandlers(bool allowed)
        {
            var responses = new PropertyBagResponseProvider("/sites/progress", allowed, "/sites/progress/subsite");
            using var context = CreateContext(responses);
            var options = new ProvisioningTemplateApplyingInformation
            {
                HandlersToProcess = Handlers.PropertyBagEntries,
                PersistTemplateInfo = false
            };

            Assert.AreEqual(allowed ? 2 : 1, InitializeTemplate(context.Web, options), "Progress must count only handlers that can run.");
            Assert.AreEqual(1, responses.ProbeWrites, "Detection must precede the first progress notification.");
            Assert.AreEqual(allowed ? 1 : 0, responses.ProbeCleanups);
            Assert.AreEqual(0, responses.NonRootProbeRequests, "Standalone subweb provisioning must probe the root web.");
            Assert.IsNull(options.PropertyBagWriteAllowed);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void HierarchyDetectsEachCollectionOnceAndReusesResultForTemplatesAndSubwebs(bool allowed)
        {
            var firstResponses = new PropertyBagResponseProvider("/sites/first", allowed, includeSubweb: true);
            var firstSubwebResponses = new PropertyBagResponseProvider("/sites/first", allowed, "/sites/first/subsite");
            var secondResponses = new PropertyBagResponseProvider("/sites/second", !allowed, includeSubweb: true);
            var secondSubwebResponses = new PropertyBagResponseProvider("/sites/second", !allowed, "/sites/second/subsite");
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
                    Sites = { new TeamNoGroupSubSite { Url = "subsite", Templates = { template.Id } } }
                });
            }

            var completedUrls = new List<string>();
            var configuration = new ApplyConfiguration();
            configuration.Handlers.Add(ConfigurationHandler.PropertyBagEntries);
            configuration.SiteProvisionedDelegate = (title, url) => completedUrls.Add(url);
            var tenant = new Tenant(context);
            var parser = new TokenParser(tenant, hierarchy);

            // Exercise the real loop: the test must not create the per-collection applying options itself.
            new ObjectHierarchySequenceSites().ProvisionObjects(tenant, hierarchy, sequence.ID, parser, configuration);

            CollectionAssert.AreEqual(new[]
            {
                firstResponses.Url, firstResponses.Url, firstSubwebResponses.Url,
                secondResponses.Url, secondResponses.Url, secondSubwebResponses.Url
            }, completedUrls);
            Assert.AreEqual(1, firstResponses.ProbeWrites);
            Assert.AreEqual(1, secondResponses.ProbeWrites, "The next collection must detect its own result.");
            Assert.AreEqual(0, firstResponses.NonRootProbeRequests + secondResponses.NonRootProbeRequests);
            Assert.AreEqual(0, firstSubwebResponses.ProbeWrites + secondSubwebResponses.ProbeWrites);
            Assert.AreEqual(allowed ? 2 : 0, firstResponses.TemplateWrites);
            Assert.AreEqual(allowed ? 1 : 0, firstSubwebResponses.TemplateWrites);
            Assert.AreEqual(allowed ? 0 : 2, secondResponses.TemplateWrites);
            Assert.AreEqual(allowed ? 0 : 1, secondSubwebResponses.TemplateWrites);
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

            public string Url { get; }
            public int ProbeWrites { get; private set; }
            public int ProbeCleanups { get; private set; }
            public int NonRootProbeRequests { get; private set; }
            public int TemplateWrites { get; private set; }

            public PropertyBagResponseProvider(string path, bool writesAllowed, string webPath = null, bool includeSubweb = false)
            {
                Url = TenantUrl + (webPath ?? path);
                this.writesAllowed = writesAllowed;
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "Web",
                    ReturnValue = CreateWebResponse(webPath ?? path, includeSubweb)
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
                        Url = TenantUrl + path,
                        EffectiveBasePermissions = new { _ObjectType_ = "SP.BasePermissions", High = 0, Low = 0 }
                    }
                });
                reads.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = Url,
                    PropertyName = "AllProperties",
                    ReturnValue = new { _ObjectType_ = "SP.PropertyValues", ExistingKey = "ExistingValue" }
                });
            }

            private static object CreateWebResponse(string path, bool includeSubweb, bool isSubweb = false)
            {
                return new
                {
                    _ObjectType_ = "SP.Web",
                    _ObjectIdentity_ = isSubweb ? "SubWeb" : "Web",
                    Url = TenantUrl + path,
                    ServerRelativeUrl = path,
                    Title = path,
                    Language = 1033,
                    WebTemplate = "STS",
                    Configuration = 0,
                    EffectiveBasePermissions = new { _ObjectType_ = "SP.BasePermissions", High = 0, Low = 0 },
                    Webs = new { _Child_Items_ = includeSubweb ? new[] { CreateWebResponse(path + "/subsite", false, true) } : Array.Empty<object>() }
                };
            }

            public string GetResponse(string url, string verb, string body)
            {
                if (body.Contains("Name=\"SetFieldValue\""))
                {
                    if (body.Contains("_PnP_PropertyBagProbe_"))
                    {
                        if (!body.Contains("Name=\"RootWeb\""))
                        {
                            NonRootProbeRequests++;
                        }
                        if (body.Contains("Type=\"Null\""))
                        {
                            ProbeCleanups++;
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

                    if (!writesAllowed)
                    {
                        return "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\"," +
                            "\"ErrorInfo\":{\"ErrorMessage\":\"Access denied\",\"ErrorCode\":-2147024891," +
                            "\"ErrorTypeName\":\"System.UnauthorizedAccessException\"}}]";
                    }

                    return "[{\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.0.0\",\"ErrorInfo\":null}]";
                }

                return reads.GetResponse(url, verb, body);
            }
        }
    }
}
