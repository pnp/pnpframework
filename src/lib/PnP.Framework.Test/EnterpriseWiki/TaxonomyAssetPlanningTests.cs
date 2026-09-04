using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Taxonomy;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class TaxonomyAssetPlanningTests
    {
        private static readonly Guid TenantId = Guid.Parse("72f988bf-86f1-41af-91ab-2d7cd011db47");
        private static readonly Guid StoreId = Guid.Parse("e385fb40-52d4-4fae-9c5b-3e8ff8a5878e");

        [TestMethod]
        public void TermIdentityUsesTenantStoreSetAndTermButNoWssId()
        {
            var identity = TaxonomyAssetIdentity.Term(new TaxonomyTermSourceIdentity
            {
                TenantId = TenantId,
                TermStoreId = StoreId,
                TermSetId = Guid.Parse("4e691f0e-5ccf-4b99-a3aa-f66b03e98a37"),
                TermId = Guid.Parse("72346334-1111-2222-3333-444444444444")
            });

            Assert.AreEqual(
                "urn:pnp:spo-term:v1:72f988bf86f141af91ab2d7cd011db47:e385fb4052d44fae9c5b3e8ff8a5878e:4e691f0e5ccf4b99a3aaf66b03e98a37:72346334111122223333444444444444",
                identity);
            Assert.IsFalse(identity.Contains("wssid", StringComparison.OrdinalIgnoreCase));
        }

        [TestMethod]
        public void RequirementCollectorDeduplicatesSharedSetsAndIncludesAnchorAndLiveOutsideTerms()
        {
            var boundSet = Guid.NewGuid();
            var liveSet = Guid.NewGuid();
            var listSet = Guid.NewGuid();
            var anchor = Guid.NewGuid();
            var liveTerm = Guid.NewGuid();
            var listTerm = Guid.NewGuid();
            var listFieldId = Guid.NewGuid();
            var list = new ListDependencySnapshot
            {
                SourceWebId = Guid.NewGuid(),
                SourceListId = Guid.NewGuid(),
                Fields = new List<ListFieldSnapshot>
                {
                    new ListFieldSnapshot
                    {
                        Id = listFieldId,
                        InternalName = "Tags",
                        Taxonomy = new TaxonomyFieldBindingSnapshot
                        {
                            SourceTermStoreId = StoreId,
                            SourceTermSetId = listSet,
                            AnchorTermId = anchor
                        }
                    }
                },
                Items = new List<ListItemSnapshot>
                {
                    new ListItemSnapshot
                    {
                        Values = new List<ListItemValueSnapshot>
                        {
                            new ListItemValueSnapshot
                            {
                                InternalName = "Tags",
                                Kind = ListItemValueKind.Taxonomy,
                                TaxonomyValues = new List<ListItemTaxonomyValueSnapshot>
                                {
                                    new ListItemTaxonomyValueSnapshot { TermGuid = listTerm.ToString("D"), WssId = 42 }
                                }
                            }
                        }
                    }
                }
            };
            var snapshot = new PublishingPageCaptureBundle
            {
                Fields = new List<PageFieldValueSnapshot>
                {
                    new PageFieldValueSnapshot
                    {
                        Id = Guid.NewGuid(),
                        InternalName = "Categories",
                        TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
                        {
                            TermStoreId = StoreId,
                            BoundTermSetId = boundSet,
                            AnchorTermId = anchor
                        },
                        TaxonomyValues = new List<PageTaxonomyValueSnapshot>
                        {
                            new PageTaxonomyValueSnapshot
                            {
                                TermGuid = liveTerm.ToString("D"),
                                WssId = 19,
                                Relationship = new TaxonomyValueRelationshipSnapshot
                                {
                                    State = TaxonomyRelationshipState.LiveOutsideBoundTermSet,
                                    LiveTermSetId = liveSet
                                }
                            }
                        }
                    }
                },
                ListDependencies = new List<ListDependencySnapshot> { list }
            };

            var requests = PublishingPageTaxonomyAssetRequirementCollector.Collect(new[] { snapshot });

            Assert.AreEqual(3, requests.Count);
            CollectionAssert.AreEquivalent(new[] { anchor }, requests.Single(value => value.SourceTermSetId == boundSet).RequiredTermIds.ToArray());
            CollectionAssert.AreEquivalent(new[] { liveTerm }, requests.Single(value => value.SourceTermSetId == liveSet).RequiredTermIds.ToArray());
            CollectionAssert.AreEquivalent(new[] { anchor, listTerm }, requests.Single(value => value.SourceTermSetId == listSet).RequiredTermIds.ToArray());
        }

        [TestMethod]
        public void RequirementCollectorIncludesValuesFromEveryCohortConsumerOfAnAffectedSet()
        {
            var setId = Guid.NewGuid();
            var termId = Guid.NewGuid();
            var gapConsumer = new PublishingPageCaptureBundle
            {
                Fields = new List<PageFieldValueSnapshot>
                {
                    new PageFieldValueSnapshot
                    {
                        Id = Guid.NewGuid(),
                        InternalName = "Methodology",
                        TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
                        {
                            TermStoreId = StoreId,
                            BoundTermSetId = setId
                        }
                    }
                }
            };
            var determinedConsumer = new PublishingPageCaptureBundle
            {
                Fields = new List<PageFieldValueSnapshot>
                {
                    new PageFieldValueSnapshot
                    {
                        Id = Guid.NewGuid(),
                        InternalName = "Methodology",
                        TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
                        {
                            TermStoreId = StoreId,
                            BoundTermSetId = setId
                        },
                        TaxonomyValues = new List<PageTaxonomyValueSnapshot>
                        {
                            new PageTaxonomyValueSnapshot
                            {
                                TermGuid = termId.ToString("D"),
                                Relationship = new TaxonomyValueRelationshipSnapshot
                                {
                                    State = TaxonomyRelationshipState.LiveInBoundTermSet
                                }
                            }
                        }
                    }
                }
            };

            var request = PublishingPageTaxonomyAssetRequirementCollector
                .Collect(new[] { gapConsumer, determinedConsumer })
                .Single();

            CollectionAssert.AreEquivalent(new[] { termId }, request.RequiredTermIds.ToArray());
            Assert.AreEqual(3, request.Consumers.Count);
        }

        [TestMethod]
        public void RequirementCollectorIncludesLayoutContentTypeTaxonomyFields()
        {
            var setId = Guid.Parse("4e691f0e-5ccf-4b99-a3aa-f66b03e98a37");
            var fieldId = Guid.Parse("42387623-5ddb-4764-94ea-e9d826afa77c");
            var snapshot = new PublishingPageCaptureBundle
            {
                Layout = new PublishingPageLayoutSnapshot
                {
                    AssociatedContentTypeSchema = new ContentTypeSchemaSnapshot
                    {
                        SourceScope = "/teams/campusipkits",
                        SourceWebUrl = "https://source.sharepoint.com/teams/campusipkits",
                        ContentTypeId = "0x010100AA",
                        RequiredFieldClosure = new List<FieldSchemaSnapshot>
                        {
                            new FieldSchemaSnapshot
                            {
                                Id = fieldId,
                                InternalName = "ActivityName",
                                Taxonomy = new TaxonomyFieldBindingSnapshot
                                {
                                    SourceTermStoreId = StoreId,
                                    SourceTermSetId = setId
                                }
                            }
                        }
                    }
                }
            };

            var request = PublishingPageTaxonomyAssetRequirementCollector.Collect(new[] { snapshot }).Single();

            Assert.AreEqual(setId, request.SourceTermSetId);
            Assert.IsTrue(request.Consumers.Any(value => value.Contains("page-layout-field:")));
        }

        [TestMethod]
        public void PlannerPreservesSourceIdsAndProducesDeterministicOwnedPlans()
        {
            var setId = Guid.NewGuid();
            var termId = Guid.NewGuid();
            var reuseSourceTermId = Guid.NewGuid();
            var additionalSetId = Guid.NewGuid();
            var pinSourceSetId = Guid.NewGuid();
            var source = new TaxonomyAssetSourceSnapshot
            {
                SourceTenantId = TenantId,
                SnapshotDigest = new string('a', 64),
                TermSets = new List<TaxonomyTermSetSourceSnapshot>
                {
                    new TaxonomyTermSetSourceSnapshot
                    {
                        SourceTenantId = TenantId,
                        SourceTermStoreId = StoreId,
                        SourceTermSetId = setId,
                        Name = "Wiki Categories",
                        IsOpenForTermCreation = true,
                        IsAvailableForTagging = true,
                        EvidenceSha256 = new string('b', 64),
                        Availability = EvidenceAvailability.Captured
                    }
                },
                Terms = new List<TaxonomyTermSourceSnapshot>
                {
                    new TaxonomyTermSourceSnapshot
                    {
                        SourceTenantId = TenantId,
                        SourceTermStoreId = StoreId,
                        SourceTermSetId = setId,
                        SourceTermId = termId,
                        Name = "Proof Points",
                        Path = "Proof Points",
                        IsAvailableForTagging = true,
                        IsReused = true,
                        IsSourceTerm = false,
                        ReuseSourceTermId = reuseSourceTermId,
                        TermSetIds = new List<Guid> { additionalSetId, setId },
                        PinSourceTermSetId = pinSourceSetId,
                        EvidenceSha256 = new string('c', 64),
                        Availability = EvidenceAvailability.Captured
                    }
                }
            };
            var targetStore = Guid.NewGuid();

            var left = TaxonomyAssetPlanner.Create(source, targetStore);
            var right = TaxonomyAssetPlanner.Create(source, targetStore);

            Assert.AreEqual(1, left.TermGroups.Count);
            Assert.AreEqual(TenantId, left.TermGroups.Single().Source.TenantId);
            Assert.AreEqual(StoreId, left.TermGroups.Single().Source.TermStoreId);
            Assert.AreEqual(setId, left.TermSets.Single().PreferredTargetTermSetId);
            Assert.AreEqual(termId, left.Terms.Single().PreferredTargetTermId);
            Assert.AreEqual(true, left.Terms.Single().SourceIsReused);
            Assert.AreEqual(false, left.Terms.Single().SourceIsSourceTerm);
            Assert.AreEqual(reuseSourceTermId, left.Terms.Single().SourceReuseSourceTermId);
            CollectionAssert.AreEqual(
                new[] { setId, additionalSetId }.OrderBy(value => value).ToArray(),
                left.Terms.Single().SourceTermSetIds.ToArray());
            Assert.AreEqual(pinSourceSetId, left.Terms.Single().SourcePinSourceTermSetId);
            Assert.AreEqual(left.TermSets.Single().TargetGroupId, right.TermSets.Single().TargetGroupId);
            Assert.AreEqual(left.TermGroups.Single().PreferredTargetGroupId, left.TermSets.Single().TargetGroupId);
            Assert.AreEqual(left.PlanDigest, right.PlanDigest);
            Assert.IsTrue(left.TermSets.Single().OriginalIdentifier.Contains(setId.ToString("N"), StringComparison.Ordinal));
        }
    }
}
