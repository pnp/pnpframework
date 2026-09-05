using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Topology.Execution;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Test.Migration
{
    [TestClass]
    public class MigrationPerformanceOptimizationTests
    {
        [TestMethod]
        public void TaxonomyFreshValidationWithNoExecutableActionsDoesNotRequireTargetContext()
        {
            var blockers = PageTaxonomyRelationshipPlanner.ValidateFresh(
                null,
                new[]
                {
                    new PageFieldValueSnapshot
                    {
                        InternalName = "WikiCategories",
                        Kind = PageFieldValueKind.TaxonomyCollection,
                        HasValue = false
                    }
                },
                new[]
                {
                    new TaxonomyRelationshipAction
                    {
                        Disposition = TaxonomyRelationshipDisposition.RetainEvidenceOnly
                    }
                },
                new PagePlanningOptions());

            Assert.AreEqual(0, blockers.Count);
        }

        [TestMethod]
        public void ReusableTopologyCompletesFromOneFreshAnalysis()
        {
            var sourceSiteId = Guid.NewGuid();
            var sourceRootWebId = Guid.NewGuid();
            var sourceChildWebId = Guid.NewGuid();
            var targetSiteId = Guid.NewGuid();
            var targetRootWebId = Guid.NewGuid();
            var targetChildWebId = Guid.NewGuid();
            var root = new WebMappingPlan
            {
                Kind = TopologyNodeKind.SiteCollectionRoot,
                SourceSiteId = sourceSiteId,
                SourceWebId = sourceRootWebId,
                TargetWebUrl = "https://target.sharepoint.com/sites/root",
                TargetServerRelativeUrl = "/sites/root"
            };
            var child = new WebMappingPlan
            {
                Kind = TopologyNodeKind.ChildWeb,
                SourceSiteId = sourceSiteId,
                SourceWebId = sourceChildWebId,
                SourceParentWebId = sourceRootWebId,
                TargetWebUrl = "https://target.sharepoint.com/sites/root/child",
                TargetServerRelativeUrl = "/sites/root/child",
                TargetParentWebUrl = root.TargetWebUrl
            };
            var plan = new TopologyPlan
            {
                PlanDigest = "topology-digest",
                SiteCollections = new List<SiteCollectionMappingPlan>
                {
                    new SiteCollectionMappingPlan
                    {
                        SourceSiteId = sourceSiteId,
                        TargetSiteCollectionUrl = root.TargetWebUrl,
                        Webs = new List<WebMappingPlan> { root, child }
                    }
                }
            };
            var analysis = new TopologyTargetAnalysis
            {
                TopologyPlanDigest = plan.PlanDigest,
                SiteCollections = new List<TopologySiteTargetProbe>
                {
                    new TopologySiteTargetProbe
                    {
                        SourceSiteId = sourceSiteId,
                        TargetSiteCollectionUrl = root.TargetWebUrl,
                        Exists = true,
                        TargetSiteId = targetSiteId,
                        TargetRootWebId = targetRootWebId,
                        Disposition = TopologyMaterializationDisposition.ReuseApprovedHost,
                        Webs = new List<TopologyWebTargetProbe>
                        {
                            new TopologyWebTargetProbe
                            {
                                SourceSiteId = sourceSiteId,
                                SourceWebId = sourceRootWebId,
                                TargetWebUrl = root.TargetWebUrl,
                                Exists = true,
                                TargetSiteId = targetSiteId,
                                TargetWebId = targetRootWebId,
                                Disposition = TopologyMaterializationDisposition.ReuseApprovedHost
                            },
                            new TopologyWebTargetProbe
                            {
                                SourceSiteId = sourceSiteId,
                                SourceWebId = sourceChildWebId,
                                TargetWebUrl = child.TargetWebUrl,
                                Exists = true,
                                TargetSiteId = targetSiteId,
                                TargetWebId = targetChildWebId,
                                TargetParentWebId = targetRootWebId,
                                Disposition = TopologyMaterializationDisposition.ReuseOwned
                            }
                        }
                    }
                }
            };
            var recorder = new MigrationExecutionRecorder(Guid.NewGuid(), "plan-digest", null);
            var receipt = new TopologyMaterializationReceipt { TopologyPlanDigest = plan.PlanDigest };

            var completed = TopologyMaterializationCoordinator.TryCompleteWithoutMutation(
                plan,
                analysis,
                recorder,
                receipt);

            Assert.IsTrue(completed);
            Assert.IsTrue(receipt.FreshReadbackPassed);
            Assert.AreEqual(2, receipt.Webs.Count);
            Assert.AreEqual(2, recorder.Steps.Count);
            Assert.AreEqual(MutationOutcome.AlreadySatisfied, recorder.Steps[0].Outcome);
            Assert.AreEqual(MutationOutcome.AlreadySatisfied, recorder.Steps[1].Outcome);
        }

        [TestMethod]
        public void TopologyFastPathRejectsAPlannedCreate()
        {
            var sourceSiteId = Guid.NewGuid();
            var sourceWebId = Guid.NewGuid();
            var plan = new TopologyPlan
            {
                PlanDigest = "topology-digest",
                SiteCollections = new List<SiteCollectionMappingPlan>
                {
                    new SiteCollectionMappingPlan
                    {
                        SourceSiteId = sourceSiteId,
                        TargetSiteCollectionUrl = "https://target.sharepoint.com/sites/root",
                        Webs = new List<WebMappingPlan>
                        {
                            new WebMappingPlan
                            {
                                Kind = TopologyNodeKind.ChildWeb,
                                SourceSiteId = sourceSiteId,
                                SourceWebId = sourceWebId,
                                TargetWebUrl = "https://target.sharepoint.com/sites/root/new",
                                TargetServerRelativeUrl = "/sites/root/new"
                            }
                        }
                    }
                }
            };
            var analysis = new TopologyTargetAnalysis
            {
                TopologyPlanDigest = plan.PlanDigest,
                SiteCollections = new List<TopologySiteTargetProbe>
                {
                    new TopologySiteTargetProbe
                    {
                        SourceSiteId = sourceSiteId,
                        Webs = new List<TopologyWebTargetProbe>
                        {
                            new TopologyWebTargetProbe
                            {
                                SourceSiteId = sourceSiteId,
                                SourceWebId = sourceWebId,
                                TargetWebUrl = "https://target.sharepoint.com/sites/root/new",
                                TargetSiteId = Guid.NewGuid(),
                                Disposition = TopologyMaterializationDisposition.CreateOwned
                            }
                        }
                    }
                }
            };

            var completed = TopologyMaterializationCoordinator.TryCompleteWithoutMutation(
                plan,
                analysis,
                new MigrationExecutionRecorder(Guid.NewGuid(), "plan-digest", null),
                new TopologyMaterializationReceipt());

            Assert.IsFalse(completed);
        }
    }
}
