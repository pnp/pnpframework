using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Execution;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ListViewRenderingResourceMigrationTests
    {
        [TestMethod]
        public void SiteCollectionTokenResolvesToTheSourceRootWeb()
        {
            var resolved = ListViewRenderingResourceSnapshotReader.ResolveSourceUri(
                new Uri("https://source.sharepoint.com/teams/source/child/"),
                new Uri("https://source.sharepoint.com/teams/source/"),
                "~sitecollection/SiteAssets/Scripts/custom.js");

            Assert.AreEqual(
                "https://source.sharepoint.com/teams/source/SiteAssets/Scripts/custom.js",
                resolved.AbsoluteUri);
        }

        [TestMethod]
        public void CapturedCustomJsLinkProducesAnExactPathResourcePlan()
        {
            var source = SourceList();
            var bytes = new byte[] { 1, 2, 3, 4 };
            var resource = new ListViewRenderingResourceSnapshot
            {
                Id = "resource-1",
                Kind = ListViewRenderingResourceKind.JavaScript,
                SourceAbsoluteUrl = "https://source.sharepoint.com/teams/source/SiteAssets/Scripts/custom.js",
                SourceServerRelativeUrl = "/teams/source/SiteAssets/Scripts/custom.js",
                Artifact = MigrationArtifact.Describe(bytes, "application/javascript", "custom.js"),
                ContentBase64 = Convert.ToBase64String(bytes),
                Availability = EvidenceAvailability.Captured
            };
            source.ViewRenderingResources.Add(resource);
            source.Views.Add(View(resource.Id));

            var planSet = ListMigrationPlanFactory.Create(
                new[] { source },
                null,
                Topology(source.SourceSiteId, source.SourceWebId),
                null,
                null);
            var plan = planSet.Lists.Single();
            var resourcePlan = plan.ViewRenderingResources.Single();

            Assert.IsFalse(plan.Issues.Any(value => value.Code == "ViewRenderingResourceUnavailable"));
            Assert.AreEqual(ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView, plan.Views.Single().Disposition);
            Assert.AreEqual(ListViewRenderingResourceMaterializationDisposition.CreateOrReuseExact, resourcePlan.Disposition);
            Assert.AreEqual("/teams/source-pnp/SiteAssets/Scripts/custom.js", resourcePlan.TargetServerRelativeUrl);
            Assert.AreEqual(resource.Artifact.Sha256, resourcePlan.SourceArtifact.Sha256);
        }

        [TestMethod]
        public void LegacyCustomJsLinkWithoutResourceClosureRemainsADeferGap()
        {
            var source = SourceList();
            source.Views.Add(View(null));

            var plan = ListMigrationPlanFactory.Create(
                new[] { source },
                null,
                Topology(source.SourceSiteId, source.SourceWebId),
                null,
                null).Lists.Single();

            Assert.AreEqual(ListViewMaterializationDisposition.Block, plan.Views.Single().Disposition);
            Assert.IsTrue(plan.Issues.Any(value => value.Code == "ViewRenderingResourceUnavailable"));
        }

        [TestMethod]
        public void UnreadableCustomJsResourcePreservesTheRelationshipWithoutBlockingTheView()
        {
            var source = SourceList();
            var resource = new ListViewRenderingResourceSnapshot
            {
                Id = "resource-1",
                Kind = ListViewRenderingResourceKind.JavaScript,
                SourceAbsoluteUrl = "https://source.sharepoint.com/teams/source/SiteAssets/Scripts/custom.js",
                SourceServerRelativeUrl = "/teams/source/SiteAssets/Scripts/custom.js",
                Availability = EvidenceAvailability.Unavailable,
                Diagnostics = { "REST file-content fallback returned HTTP 404." }
            };
            source.ViewRenderingResources.Add(resource);
            source.Views.Add(View(resource.Id));

            var plan = ListMigrationPlanFactory.Create(
                new[] { source },
                null,
                Topology(source.SourceSiteId, source.SourceWebId),
                null,
                null).Lists.Single();

            Assert.AreEqual(ListViewRenderingResourceMaterializationDisposition.PreserveReferenceOnly, plan.ViewRenderingResources.Single().Disposition);
            Assert.AreEqual(ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView, plan.Views.Single().Disposition);
            Assert.IsFalse(plan.Issues.Any(value => value.Code == "ViewRenderingResourceUnavailable"));
        }

        [TestMethod]
        public void SiteCollectionTokenIsPreservedAfterItsBytesMoveToTheMappedSite()
        {
            var view = View("resource-1");
            var rewritten = ListViewRenderingResourceMaterializer.RewriteJsLink(
                view.JsLink,
                view,
                new[]
                {
                    new ListViewRenderingResourceMaterializationPlan
                    {
                        SourceResourceId = "resource-1",
                        TargetAbsoluteUrl = "https://target.sharepoint.com/teams/source-pnp/SiteAssets/Scripts/custom.js",
                        TargetServerRelativeUrl = "/teams/source-pnp/SiteAssets/Scripts/custom.js",
                        Disposition = ListViewRenderingResourceMaterializationDisposition.CreateOrReuseExact
                    }
                });

            Assert.AreEqual(
                "clienttemplates.js|~sitecollection/SiteAssets/Scripts/custom.js",
                rewritten);
        }

        private static ListDependencySnapshot SourceList()
        {
            return new ListDependencySnapshot
            {
                SourceSiteId = Guid.Parse("11111111-1111-1111-1111-111111111111"),
                SourceWebId = Guid.Parse("22222222-2222-2222-2222-222222222222"),
                SourceWebUrl = "https://source.sharepoint.com/teams/source",
                SourceListId = Guid.Parse("33333333-3333-3333-3333-333333333333"),
                Title = "Documents",
                BaseTemplate = 101,
                BaseType = "DocumentLibrary",
                RootFolderServerRelativeUrl = "/teams/source/Documents",
                Availability = EvidenceAvailability.Captured
            };
        }

        private static ListViewSnapshot View(string resourceId)
        {
            var view = new ListViewSnapshot
            {
                Id = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                Title = "All Documents",
                ServerRelativeUrl = "/teams/source/Documents/Forms/AllItems.aspx",
                ViewType = "Html",
                ListViewXml = "<View />",
                ListViewXmlSha256 = MigrationDigest.ComputeSha256("<View />"),
                JsLink = "clienttemplates.js|~sitecollection/SiteAssets/Scripts/custom.js",
                Availability = EvidenceAvailability.Captured
            };
            if (!string.IsNullOrWhiteSpace(resourceId))
            {
                view.RenderingResourceBindings.Add(new ListViewRenderingResourceBindingSnapshot
                {
                    SourceProperty = "JSLink",
                    OriginalReference = "~sitecollection/SiteAssets/Scripts/custom.js",
                    ResourceId = resourceId
                });
            }
            return view;
        }

        private static TopologyPlan Topology(Guid siteId, Guid webId)
        {
            return new TopologyPlan
            {
                SiteCollections = new List<SiteCollectionMappingPlan>
                {
                    new SiteCollectionMappingPlan
                    {
                        SourceSiteId = siteId,
                        SourceSiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                        TargetSiteCollectionUrl = "https://target.sharepoint.com/teams/source-pnp",
                        Webs = new List<WebMappingPlan>
                        {
                            new WebMappingPlan
                            {
                                Kind = TopologyNodeKind.SiteCollectionRoot,
                                SourceSiteId = siteId,
                                SourceWebId = webId,
                                SourceSiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                                SourceWebUrl = "https://source.sharepoint.com/teams/source",
                                SourceServerRelativeUrl = "/teams/source",
                                TargetSiteCollectionUrl = "https://target.sharepoint.com/teams/source-pnp",
                                TargetWebUrl = "https://target.sharepoint.com/teams/source-pnp",
                                TargetServerRelativeUrl = "/teams/source-pnp"
                            }
                        }
                    }
                }
            };
        }
    }
}
