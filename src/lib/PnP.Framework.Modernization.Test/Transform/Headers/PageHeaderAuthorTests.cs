using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Publishing;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Modernization.Tests.Transform.Headers
{
    [TestClass]
    [DoNotParallelize]
    public class PageHeaderAuthorTests
    {
        [DataTestMethod]
        [DataRow(null, false)]
        [DataRow(false, false)]
        [DataRow(true, false)]
        [DataRow(null, true)]
        public void WikiAndWebPartDefaultHeader_IgnoresCompatibilityFlag(bool? flag, bool explicitNone)
        {
            using (var fixture = new Fixture())
            {
                var information = fixture.WikiInformation();
                if (flag.HasValue)
                {
                    information.SetAuthorInPageHeader = flag.Value;
                }
                if (explicitNone)
                {
                    information.PageHeader = Mock.Of<PnPCore.IPageHeader>(h => h.Type == PnPCore.PageHeaderType.None);
                }
                var transformator = fixture.WikiTransformator(information);

                transformator.ConfigurePageHeader(information, fixture.Page.Object, fixture.Target, alwaysSetAuthor: true);

                fixture.Page.Verify(p => p.SetDefaultPageHeader(), Times.Once);
                fixture.Page.Verify(p => p.RemovePageHeader(), Times.Never);
                Assert.AreEqual(PnPCore.PageHeaderLayoutType.NoImage, fixture.Header.Object.LayoutType);
                fixture.AssertAuthor("author@contoso.com", 42, "Target Author", "Target role");
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Never);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void BlogHeader_RetainsExistingOptIn(bool flag)
        {
            using (var fixture = new Fixture())
            {
                var information = fixture.WikiInformation();
                information.SetAuthorInPageHeader = flag;
                fixture.WikiTransformator(information)
                    .ConfigurePageHeader(information, fixture.Page.Object, fixture.Target, alwaysSetAuthor: false);
                fixture.Page.Verify(p => p.RemovePageHeader(), flag ? Times.Never() : Times.Once());
                fixture.Page.Verify(p => p.SetDefaultPageHeader(), flag ? Times.Once() : Times.Never());
            }
        }

        [DataTestMethod]
        [DataRow(PnPCore.PageHeaderType.Default)]
        [DataRow(PnPCore.PageHeaderType.Custom)]
        public void ExplicitDotNetHeader_RetainsExistingAppearance(PnPCore.PageHeaderType type)
        {
            using (var fixture = new Fixture())
            {
                var information = fixture.WikiInformation();
                information.PageHeader = Mock.Of<PnPCore.IPageHeader>(h => h.Type == type &&
                    h.ImageServerRelativeUrl == "/sites/target/banner.png" && h.TranslateX == 10 && h.TranslateY == 20);

                fixture.WikiTransformator(information)
                    .ConfigurePageHeader(information, fixture.Page.Object, fixture.Target, alwaysSetAuthor: true);

                if (type == PnPCore.PageHeaderType.Custom)
                {
                    fixture.Page.Verify(p => p.SetCustomPageHeader("/sites/target/banner.png", 10, 20), Times.Once);
                }
                else
                {
                    fixture.Page.Verify(p => p.SetDefaultPageHeader(), Times.Once);
                }
                Assert.IsNull(fixture.Header.Object.Authors);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void HomePage_OnlyStockReplacementRetainsHomeLayout(bool replace)
        {
            var page = new Mock<PnPCore.IPage>();
            page.SetupAllProperties();

            var skipConversion = PageTransformator.ConfigureHomePage(page.Object, replace);

            Assert.AreEqual(replace, skipConversion);
            Assert.AreEqual(replace, page.Object.KeepDefaultWebParts);
            Assert.AreEqual(replace ? PnPCore.PageLayoutType.Home : PnPCore.PageLayoutType.Article, page.Object.LayoutType);
            page.Verify(p => p.SetDefaultPageHeader(), Times.Never);
        }

        [DataTestMethod]
        [DataRow(PageLayoutPageHeader.None, false)]
        [DataRow(PageLayoutPageHeader.Default, false)]
        [DataRow(PageLayoutPageHeader.Default, true)]
        public void PublishingNoneAndDefault_UseCreatedByIndependentlyOfAuditFlag(PageLayoutPageHeader type, bool keepAudit)
        {
            using (var fixture = new Fixture())
            {
                var information = fixture.PublishingInformation();
                information.KeepPageCreationModificationInformation = keepAudit;
                var transformator = fixture.PublishingTransformator(information, type);
                var page = fixture.Page.Object;

                transformator.TransformHeader(ref page);

                Assert.AreEqual(PnPCore.PageHeaderLayoutType.NoImage, fixture.Header.Object.LayoutType);
                fixture.AssertAuthor("author@contoso.com", 42, "Target Author", "Target role");
                Assert.AreEqual(7, ((FieldUserValue)fixture.SourcePage[Constants.CreatedByField]).LookupId);
            }
        }

        [TestMethod]
        public void PublishingCustom_UsesContactAndRetainsMappedHeaderSettings()
        {
            using (var fixture = new Fixture())
            {
                var transformator = fixture.PublishingTransformator(fixture.PublishingInformation(), PageLayoutPageHeader.Custom);
                var page = fixture.Page.Object;

                transformator.TransformHeader(ref page);

                fixture.AssertAuthor("contact@contoso.com", 43, "Contact", "Contact role");
                Assert.AreEqual(PnPCore.PageHeaderLayoutType.ColorBlock, fixture.Header.Object.LayoutType);
                Assert.AreEqual(PnPCore.PageHeaderTitleAlignment.Center, fixture.Header.Object.TextAlignment);
                Assert.IsTrue(fixture.Header.Object.ShowPublishDate);
                Assert.IsTrue(fixture.Header.Object.ShowTopicHeader);
                Assert.AreEqual("Article topic", fixture.Header.Object.TopicHeader);
                Assert.AreEqual("Image description", fixture.Header.Object.AlternativeText);
            }
        }

        [TestMethod]
        public void PublishingCustom_EmptyContactDoesNotFallBackToCreatedBy()
        {
            using (var fixture = new Fixture())
            {
                fixture.SourcePage["PublishingContact"] = null;
                var transformator = fixture.PublishingTransformator(fixture.PublishingInformation(), PageLayoutPageHeader.Custom);
                var page = fixture.Page.Object;

                transformator.TransformHeader(ref page);

                Assert.IsNull(fixture.Header.Object.Authors);
                Assert.AreEqual(-1, fixture.Header.Object.AuthorByLineId);
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Once);
            }
        }

        [TestMethod]
        public void CreatedBy_UsesMappedTargetIdentityAndJobTitle()
        {
            using (var fixture = new Fixture())
            {
                fixture.StoreMappings(new List<UserMappingEntity>
                {
                    new UserMappingEntity { SourceUser = "author@contoso.com", TargetUser = "contact@contoso.com" }
                });
                var information = fixture.WikiInformation();
                information.UserMappingFile = "cached-mapping.csv";
                fixture.WikiTransformator(information)
                    .ConfigurePageHeader(information, fixture.Page.Object, fixture.Target, alwaysSetAuthor: true);

                fixture.AssertAuthor("contact@contoso.com", 43, "Contact", "Contact role");
            }
        }

        [TestMethod]
        public void SameSiteHeader_DoesNotRequireCrossSiteUserMappingInitialization()
        {
            using (var fixture = new Fixture(sameSite: true))
            {
                var information = fixture.WikiInformation();
                var transformator = new PageTransformator(fixture.Source, new PageTransformation());
                Assert.IsNull(transformator.userTransformator);

                transformator.ConfigurePageHeader(information, fixture.Page.Object, fixture.Source, alwaysSetAuthor: true);

                fixture.AssertAuthor("author@contoso.com", 7, "Target Author", "Target role");
            }
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("Lead \"Ing\u00e9nieur\" & Architect")]
        public void CreatedBy_JobTitleIsPreservedAsJson(string role)
        {
            using (var fixture = new Fixture())
            {
                fixture.SeedUser(fixture.Target, 42, "author@contoso.com", "Target Author", role);
                fixture.AuthorTransformator().SetSourceAuthor(fixture.Header.Object, fixture.SourcePage);
                fixture.AssertAuthor("author@contoso.com", 42, "Target Author", role);
            }
        }

        [DataTestMethod]
        [DataRow("missing-field")]
        [DataRow("null-page")]
        [DataRow("unresolved-source")]
        [DataRow("source-group")]
        [DataRow("unresolved-target")]
        [DataRow("target-group")]
        [DataRow("system-account")]
        [DataRow("application-account")]
        public void UnavailableAuthor_KeepsNoImageHeaderAndWarns(string failure)
        {
            using (var fixture = new Fixture())
            {
                var information = fixture.WikiInformation();
                if (failure == "missing-field") fixture.SourcePage[Constants.CreatedByField] = null;
                if (failure == "null-page") information.SourcePage = null;
                if (failure == "unresolved-source") fixture.CacheUser(fixture.Source, 7, null);
                if (failure == "source-group") fixture.SeedUser(fixture.Source, 7, "author@contoso.com", "Group", "", isGroup: true);
                if (failure == "unresolved-target") fixture.CacheEnsuredUser("author@contoso.com", null);
                if (failure == "target-group") fixture.SeedUser(fixture.Target, 42, "author@contoso.com", "Group", "", isGroup: true);
                if (failure == "system-account") fixture.CacheEnsuredUser("author@contoso.com",
                    new ResolvedUser { Id = 42, LoginName = @"SHAREPOINT\system" });
                if (failure == "application-account") fixture.CacheEnsuredUser("author@contoso.com",
                    new ResolvedUser { Id = 42, LoginName = "app@sharepoint" });

                fixture.WikiTransformator(information)
                    .ConfigurePageHeader(information, fixture.Page.Object, fixture.Target, alwaysSetAuthor: true);

                Assert.AreEqual(PnPCore.PageHeaderLayoutType.NoImage, fixture.Header.Object.LayoutType);
                Assert.IsNull(fixture.Header.Object.Authors);
                Assert.AreEqual(-1, fixture.Header.Object.AuthorByLineId);
                Assert.AreEqual("Page title", fixture.Page.Object.PageTitle);
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Once);
            }
        }

        [DataTestMethod]
        [DataRow(false, "upn")]
        [DataRow(true, "email")]
        [DataRow(true, "id")]
        public void CustomAuthorJson_PreservesDisplayDataAndCompletesByline(bool array, string identityProperty)
        {
            using (var fixture = new Fixture())
            {
                var author = new JObject
                {
                    [identityProperty] = identityProperty == "id" ? "i:0#.f|membership|contact@contoso.com" : "contact@contoso.com",
                    ["name"] = "Custom display name",
                    ["role"] = "Custom role",
                    ["extension"] = new JObject { ["keep"] = true }
                };
                var json = array ? new JArray(author).ToString() : author.ToString();

                fixture.AuthorTransformator().SetMappedAuthors(fixture.Header.Object, json);

                Assert.AreEqual(json, fixture.Header.Object.Authors);
                Assert.AreEqual("contact@contoso.com", JArray.Parse(fixture.Header.Object.AuthorByLine)[0].Value<string>());
                Assert.AreEqual(43, fixture.Header.Object.AuthorByLineId);
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Never);
            }
        }

        [DataTestMethod]
        [DataRow("{\"name\":\"Display only\"}")]
        [DataRow("[{\"upn\":\"author@contoso.com\"},{\"upn\":\"contact@contoso.com\"}]")]
        [DataRow("{invalid}")]
        [DataRow("\"Display name\"")]
        [DataRow("{\"upn\":123}")]
        public void UnsupportedCustomAuthor_PreservesJsonAndWarnsWithoutGuessing(string json)
        {
            using (var fixture = new Fixture())
            {
                fixture.AuthorTransformator().SetMappedAuthors(fixture.Header.Object, json);
                Assert.AreEqual(json, fixture.Header.Object.Authors);
                Assert.IsNull(fixture.Header.Object.AuthorByLine);
                Assert.AreEqual(-1, fixture.Header.Object.AuthorByLineId);
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Once);
            }
        }

        [TestMethod]
        public void EmptyCustomAuthorArray_DoesNotInventAnAuthor()
        {
            using (var fixture = new Fixture())
            {
                fixture.AuthorTransformator().SetMappedAuthors(fixture.Header.Object, "[]");
                Assert.AreEqual("[]", fixture.Header.Object.Authors);
                Assert.IsNull(fixture.Header.Object.AuthorByLine);
                fixture.Observer.Verify(o => o.Warning(It.IsAny<LogEntry>()), Times.Never);
            }
        }

        [TestMethod]
        public void UserLookup_UsesExactNumericId()
        {
            using (var fixture = new Fixture())
            {
                fixture.Provider.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = fixture.Source.Url,
                    Method = "GetItems",
                    ReturnValue = new
                    {
                        _Child_Items_ = new[]
                        {
                            new { _ObjectType_ = "SP.ListItem", ID = 17, Title = "Exact user",
                                Name = "i:0#.f|membership|exact@contoso.com", JobTitle = "Engineer" }
                        }
                    }
                });
                string requestBody = null;
                var provider = new Mock<IMockResponseProvider>();
                provider.Setup(p => p.GetResponse(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                    .Returns<string, string, string>((url, verb, body) =>
                    {
                        requestBody = body;
                        return fixture.Provider.GetResponse(url, verb, body);
                    });
                fixture.Source.WebRequestExecutorFactory = new MockWebRequestExecutorFactory(provider.Object);

                var author = CacheManager.Instance.GetUserFromUserList(fixture.Source, 17);

                Assert.IsNotNull(author);
                Assert.AreEqual("Exact user", author.Name);
                var request = XDocument.Parse(requestBody);
                var viewXml = request.Descendants().Single(e => e.Name.LocalName == "Property" && (string)e.Attribute("Name") == "ViewXml").Value;
                var predicate = XDocument.Parse(viewXml).Descendants("Where").Single().Elements().Single();
                Assert.AreEqual("Eq", predicate.Name.LocalName);
                Assert.AreEqual("ID", (string)predicate.Element("FieldRef").Attribute("Name"));
                Assert.AreEqual("17", predicate.Element("Value").Value);
            }
        }

        private sealed class Fixture : IDisposable
        {
            internal readonly ClientContext Source;
            internal readonly ClientContext Target;
            internal readonly ListItem SourcePage;
            internal readonly MockEntryResponseProvider Provider = new MockEntryResponseProvider();
            internal readonly Mock<ILogObserver> Observer = new Mock<ILogObserver>();
            internal readonly Mock<PnPCore.IPage> Page = new Mock<PnPCore.IPage>();
            internal readonly Mock<PnPCore.IPageHeader> Header = new Mock<PnPCore.IPageHeader>();
            private byte[] savedMappings;
            private bool changedMappings;

            internal Fixture(bool sameSite = false)
            {
                var suffix = Guid.NewGuid().ToString("N");
                Source = CreateContext("https://contoso.sharepoint.com/sites/source-" + suffix);
                Target = sameSite ? Source : CreateContext("https://contoso.sharepoint.com/sites/target-" + suffix);
                SourcePage = Source.Web.Lists.GetByTitle("Pages").GetItemById(1);
                SourcePage[Constants.CreatedByField] = new FieldUserValue { LookupId = 7 };
                SourcePage[Constants.PublishingPageLayoutField] = new FieldUrlValue { Url = "https://contoso.sharepoint.com/_catalogs/masterpage/Test.aspx" };
                SourcePage["PublishingContact"] = new FieldUserValue { LookupId = 13 };
                SourcePage["ArticleByLine"] = "Article topic";
                SourcePage["ImageAlt"] = "Image description";
                SeedUser(Source, 7, "author@contoso.com", "Source Author", "Source role");
                SeedUser(Source, 13, "contact@contoso.com", "Contact", "Contact role");
                SeedUser(Target, sameSite ? 7 : 42, "author@contoso.com", "Target Author", "Target role");
                SeedUser(Target, sameSite ? 13 : 43, "contact@contoso.com", "Contact", "Contact role");
                CacheEnsuredUser("author@contoso.com", new ResolvedUser { Id = sameSite ? 7 : 42, LoginName = "i:0#.f|membership|author@contoso.com" });
                CacheEnsuredUser("contact@contoso.com", new ResolvedUser { Id = sameSite ? 13 : 43, LoginName = "i:0#.f|membership|contact@contoso.com" });
                Header.SetupAllProperties();
                Header.Object.AuthorByLineId = -1;
                Page.SetupAllProperties();
                Page.SetupGet(p => p.PageHeader).Returns(Header.Object);
                Page.Object.PageTitle = "Page title";
            }

            internal PageTransformationInformation WikiInformation() =>
                new PageTransformationInformation(SourcePage) { SourceVersion = SPVersion.SPO, SkipTelemetry = true };

            internal PublishingPageTransformationInformation PublishingInformation() =>
                new PublishingPageTransformationInformation(SourcePage) { SourceVersion = SPVersion.SPO, SkipTelemetry = true };

            internal PageTransformator WikiTransformator(PageTransformationInformation information)
            {
                var transformator = new PageTransformator(Source, Target, new PageTransformation());
                transformator.RegisterObserver(Observer.Object);
                transformator.InitializeUserMapping(information);
                return transformator;
            }

            internal PageHeaderAuthorTransformator AuthorTransformator() =>
                new PageHeaderAuthorTransformator(Source, Target, new UserTransformator(WikiInformation(), Source, Target),
                    new[] { Observer.Object });

            internal PublishingPageHeaderTransformator PublishingTransformator(PublishingPageTransformationInformation information, PageLayoutPageHeader type)
            {
                var model = new PublishingPageTransformation
                {
                    PageLayouts = new[]
                    {
                        new PnP.Framework.Modernization.Publishing.PageLayout
                        {
                            Name = "Test", PageHeader = type,
                            Header = new Header
                            {
                                Type = HeaderType.ColorBlock, Alignment = HeaderAlignment.Center, ShowPublishedDate = true,
                                Field = new[]
                                {
                                    new HeaderField { Name = "PublishingContact", HeaderProperty = HeaderFieldHeaderProperty.Authors, Functions = "ToAuthors({PublishingContact})" },
                                    new HeaderField { Name = "ArticleByLine", HeaderProperty = HeaderFieldHeaderProperty.TopicHeader },
                                    new HeaderField { Name = "ImageAlt", HeaderProperty = HeaderFieldHeaderProperty.AlternativeText }
                                }
                            }
                        }
                    }
                };
                return new PublishingPageHeaderTransformator(information, Source, Target, model, new[] { Observer.Object });
            }

            internal void AssertAuthor(string upn, int id, string name, string role)
            {
                var author = JObject.Parse(Header.Object.Authors);
                Assert.AreEqual(upn, (string)author["upn"]);
                Assert.AreEqual(name, (string)author["name"]);
                Assert.AreEqual(role, (string)author["role"]);
                Assert.AreEqual(upn, JArray.Parse(Header.Object.AuthorByLine)[0].Value<string>());
                Assert.AreEqual(id, Header.Object.AuthorByLineId);
            }

            internal void SeedUser(ClientContext context, int id, string upn, string name, string role, bool isGroup = false)
            {
                var author = new UserEntity
                {
                    Id = "i:0#.f|membership|" + upn, LoginName = "i:0#.f|membership|" + upn,
                    Upn = upn, Name = name, Role = role, IsGroup = isGroup
                };
                CacheUser(context, id, author);
                CacheEntry("userJsonStringsViaUpn", context.Url, author.LoginName, author);
            }

            internal void CacheUser(ClientContext context, int id, UserEntity author) =>
                CacheEntry("userJsonStrings", context.Url, id.ToString(), author);

            internal void CacheEnsuredUser(string upn, ResolvedUser user)
            {
                CacheEntry("ensuredUsers", Target.Url, upn, user);
                CacheEntry("ensuredUsers", Target.Url, "i:0#.f|membership|" + upn, user);
            }

            internal void StoreMappings(List<UserMappingEntity> mappings)
            {
                var cache = CacheManager.Instance;
                var key = cache.StoreOptions.GetKey("userMappings");
                savedMappings = cache.Store.Get(key);
                changedMappings = true;
                cache.Store.Set(key, mappings, cache.StoreOptions.EntryOptions);
            }

            private static void CacheEntry<T>(string key, string url, string id, T value) where T : class
            {
                var cache = CacheManager.Instance;
                var qualifiedKey = cache.StoreOptions.GetKey(key);
                var entries = cache.Store.GetAndInitialize<Dictionary<string, Dictionary<string, T>>>(qualifiedKey);
                if (!entries.TryGetValue(url, out var users))
                {
                    users = new Dictionary<string, T>();
                    entries[url] = users;
                }
                users[id] = value;
                cache.Store.Set(qualifiedKey, entries, cache.StoreOptions.EntryOptions);
            }

            private ClientContext CreateContext(string url)
            {
                Provider.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = url, PropertyName = "Web",
                    ReturnValue = new { _ObjectType_ = "SP.Web", Url = url, ServerRelativeUrl = new Uri(url).AbsolutePath }
                });
                Provider.ResponseEntries.Add(new MockResponseEntry
                {
                    Url = url, PropertyName = "Site",
                    ReturnValue = new { _ObjectType_ = "SP.Site", Url = url, ServerRelativeUrl = new Uri(url).AbsolutePath }
                });
                var context = new ClientContext(url) { WebRequestExecutorFactory = new MockWebRequestExecutorFactory(Provider) };
                context.Load(context.Web, w => w.Url);
                context.Load(context.Site, s => s.Url, s => s.ServerRelativeUrl);
                context.ExecuteQuery();
                return context;
            }

            public void Dispose()
            {
                RemoveEntries<UserEntity>("userJsonStrings");
                RemoveEntries<UserEntity>("userJsonStringsViaUpn");
                RemoveEntries<ResolvedUser>("ensuredUsers");
                if (changedMappings)
                {
                    var cache = CacheManager.Instance;
                    var key = cache.StoreOptions.GetKey("userMappings");
                    if (savedMappings == null) cache.Store.Remove(key);
                    else cache.Store.Set(key, savedMappings, cache.StoreOptions.EntryOptions);
                }
                if (!ReferenceEquals(Source, Target)) Target.Dispose();
                Source.Dispose();
            }

            private void RemoveEntries<T>(string key) where T : class
            {
                var cache = CacheManager.Instance;
                var qualifiedKey = cache.StoreOptions.GetKey(key);
                var entries = cache.Store.GetAndInitialize<Dictionary<string, Dictionary<string, T>>>(qualifiedKey);
                entries.Remove(Source.Url);
                entries.Remove(Target.Url);
                cache.Store.Set(qualifiedKey, entries, cache.StoreOptions.EntryOptions);
            }
        }
    }
}
