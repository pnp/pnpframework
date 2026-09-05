using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PlatformFeatureMigrationTests
    {
        [TestMethod]
        public void VideoContentTypePlansConditionalRuntimeFeatures()
        {
            var plan = CreateVideoListPlan();
            var features = plan.RequiredFeatures.ToDictionary(value => value.FeatureId);

            Assert.AreEqual(3, features.Count);
            Assert.IsTrue(features.ContainsKey(ContentTypeRuntimeCatalog.AssetLibraryFeatureId));
            Assert.IsTrue(features.ContainsKey(ContentTypeRuntimeCatalog.DocumentSetFeatureId));
            Assert.IsTrue(features.ContainsKey(ContentTypeRuntimeCatalog.VideoAndRichMediaFeatureId));
            CollectionAssert.AreEqual(
                new[] { BuiltInContentTypeId.DocumentSet },
                features[ContentTypeRuntimeCatalog.DocumentSetFeatureId].ExpectedContentTypeIds.ToArray());
            CollectionAssert.AreEqual(
                new[] { "0x0120D520A8", "0x0120D520A808" },
                features[ContentTypeRuntimeCatalog.VideoAndRichMediaFeatureId].ExpectedContentTypeIds.ToArray());
            Assert.IsTrue(features.Values.All(value => value.TargetWebUrl == "https://target.sharepoint.com/sites/target"));
        }

        [TestMethod]
        public void ListSemanticDigestExcludesPlatformFeatureTargetProbe()
        {
            var plan = CreateVideoListPlan();
            var expected = plan.PlanDigest;
            foreach (var feature in plan.RequiredFeatures)
            {
                feature.TargetProbe = new PlatformFeatureTargetProbe
                {
                    FeatureId = feature.FeatureId,
                    Scope = feature.Scope,
                    TargetWebUrl = feature.TargetWebUrl,
                    TargetScopeExists = true,
                    IsActive = true,
                    CanActivate = true,
                    AvailableContentTypeIds = feature.ExpectedContentTypeIds.ToList()
                };
            }

            Assert.AreEqual(expected, ListMigrationPlanFactory.ComputePlanDigest(plan));
        }

        [TestMethod]
        public void InactivePlatformFeatureIsAdmittedOnlyWhenActivationIsAuthorized()
        {
            var probe = new PlatformFeatureTargetProbe
            {
                FeatureId = ContentTypeRuntimeCatalog.VideoAndRichMediaFeatureId,
                Scope = PlatformFeatureScope.SiteCollection,
                TargetScopeExists = true,
                IsActive = false,
                CanActivate = true
            };

            Assert.IsTrue(probe.IsAdmitted);
            probe.CanActivate = false;
            Assert.IsFalse(probe.IsAdmitted);
        }

        [TestMethod]
        public void PlatformFeatureActionRemainsAvailableWhenConsumingListIsBlocked()
        {
            var source = CreateVideoListSource();
            var listPlan = CreateVideoListPlan(source);
            listPlan.Disposition = ListMaterializationDisposition.Block;
            var snapshot = new PublishingPageCaptureBundle
            {
                ListDependencies = new List<ListDependencySnapshot> { source }
            };
            var plan = new PublishingPageMigrationPlan
            {
                ListMigration = new ListMigrationPlanSet
                {
                    Lists = new List<ListMaterializationPlan> { listPlan }
                }
            };
            var actions = new Dictionary<string, PageIngredientAction>();

            PublishingPageListIngredientActionProjector.Project(snapshot, plan, actions);

            var featureAction = actions[PublishingPageIngredientIds.PlatformFeature(
                source.SourceSiteId,
                ContentTypeRuntimeCatalog.VideoAndRichMediaFeatureId)];
            Assert.AreEqual(IngredientCapability.Available, featureAction.Capability);
            Assert.AreEqual(IngredientDisposition.Substitute, featureAction.Disposition);
            Assert.AreEqual("activate-target-runtime-feature", featureAction.Realization);
        }

        private static ListMaterializationPlan CreateVideoListPlan()
        {
            return CreateVideoListPlan(CreateVideoListSource());
        }

        private static ListMaterializationPlan CreateVideoListPlan(ListDependencySnapshot source)
        {
            var topology = CreateTopology(source.SourceSiteId, source.SourceWebId);
            return ListMigrationPlanFactory.Create(new[] { source }, null, topology, null, null).Lists.Single();
        }

        private static ListDependencySnapshot CreateVideoListSource()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            return new ListDependencySnapshot
            {
                SourceSiteId = siteId,
                SourceWebId = webId,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceListId = listId,
                Title = "Assets",
                BaseTemplate = 851,
                BaseType = "DocumentLibrary",
                RootFolderServerRelativeUrl = "/sites/source/Assets",
                SourceItemCount = 0,
                Availability = EvidenceAvailability.Captured,
                ContentTypes = new List<ListContentTypeSnapshot>
                {
                    new ListContentTypeSnapshot
                    {
                        Id = "0x0120D520A80800AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                        Name = "Video",
                        ParentId = "0x0120D520A808"
                    }
                }
            };
        }

        private static TopologyPlan CreateTopology(Guid siteId, Guid webId)
        {
            var topology = new TopologyPlan
            {
                SiteCollections = new List<SiteCollectionMappingPlan>
                {
                    new SiteCollectionMappingPlan
                    {
                        SourceSiteId = siteId,
                        TargetSiteCollectionUrl = "https://target.sharepoint.com/sites/target",
                        Webs = new List<WebMappingPlan>
                        {
                            new WebMappingPlan
                            {
                                Kind = TopologyNodeKind.SiteCollectionRoot,
                                SourceSiteId = siteId,
                                SourceWebId = webId,
                                SourceServerRelativeUrl = "/sites/source",
                                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                                TargetServerRelativeUrl = "/sites/target"
                            }
                        }
                    }
                }
            };
            topology.PlanDigest = TopologyPlanner.ComputeDigest(topology);
            return topology;
        }
    }
}
