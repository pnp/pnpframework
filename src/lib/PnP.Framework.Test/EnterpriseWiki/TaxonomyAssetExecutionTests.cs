using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets;
using PnP.Framework.Migration.Taxonomy.Assets.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using PnP.Framework.Migration.Taxonomy.Assets.Verification;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class TaxonomyAssetExecutionTests
    {
        private static readonly Guid TenantId = Guid.Parse("72f988bf-86f1-41af-91ab-2d7cd011db47");
        private static readonly Guid SourceStoreId = Guid.Parse("e385fb40-52d4-4fae-9c5b-3e8ff8a5878e");
        private static readonly Guid TargetStoreId = Guid.Parse("c5e18914-52aa-4047-8ef6-f9654987b925");
        private static readonly Guid SourceSetId = Guid.Parse("787ae7d4-495e-46c2-a3be-066d33fcfced");
        private static readonly Guid SourceTermId = Guid.Parse("67984a5d-e21d-4f50-9a30-cede4c211a5e");

        [TestMethod]
        public void ApprovalTemplateAutoApprovesOnlyExactOwnedReuse()
        {
            var plan = Plan(
                TaxonomyAssetTargetDisposition.ReuseOwned,
                TaxonomyAssetTargetDisposition.ReuseOwned,
                SourceSetId);

            var template = TaxonomyAssetApprovalFactory.CreateTemplate(plan, DateTimeOffset.UtcNow);

            Assert.AreEqual(3, template.Actions.Count);
            Assert.IsTrue(template.Actions.All(value => value.Decision == TaxonomyAssetApprovalDecision.Approve));
            Assert.IsTrue(template.Actions.All(value => !value.RequiresExplicitReview));
            Assert.IsNull(template.ApprovalDigest);
        }

        [TestMethod]
        public void ExternalChildCreationRequiresSeparateMutationApproval()
        {
            var externalSetId = Guid.NewGuid();
            var plan = Plan(
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval,
                externalSetId);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(plan, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = TaxonomyAssetApprovalDecision.Approve;
            }

            Assert.ThrowsException<InvalidDataException>(() => TaxonomyAssetApprovalFactory.Seal(
                plan,
                approval,
                "reviewer@example.com",
                DateTimeOffset.UtcNow));

            approval.Actions.Single(value => value.Kind == TaxonomyAssetKind.Term).ExternalMutationApproved = true;
            TaxonomyAssetApprovalFactory.Seal(plan, approval, "reviewer@example.com", DateTimeOffset.UtcNow);

            Assert.AreEqual(64, approval.ApprovalDigest.Length);
            TaxonomyAssetApprovalValidator.Validate(plan, approval);
        }

        [TestMethod]
        public void AdmissionAcceptsOwnedAssetsCreatedByAnInterruptedPriorAttempt()
        {
            var reviewed = Plan(
                TaxonomyAssetTargetDisposition.CreateMissing,
                TaxonomyAssetTargetDisposition.CreateMissing,
                SourceSetId,
                TaxonomyAssetTargetDisposition.CreateMissing);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(reviewed, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = TaxonomyAssetApprovalDecision.Approve;
            }
            TaxonomyAssetApprovalFactory.Seal(reviewed, approval, "reviewer@example.com", DateTimeOffset.UtcNow);

            var fresh = TaxonomyAssetContractCloner.Clone(reviewed);
            fresh.TermGroupProbes.Single().Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
            fresh.TermGroupProbes.Single().ResolvedTargetGroupId = reviewed.TermGroups.Single().PreferredTargetGroupId;
            fresh.TermSetProbes.Single().Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
            fresh.TermSetProbes.Single().ResolvedTargetTermSetId = SourceSetId;
            fresh.TermProbes.Single().Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
            fresh.TermProbes.Single().ResolvedTargetTermId = SourceTermId;
            fresh.MappingCandidates.Single().Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
            fresh.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(fresh);

            var admission = TaxonomyAssetExecutionAdmissionEvaluator.Evaluate(reviewed, fresh, approval);

            Assert.IsTrue(admission.IsAdmitted);
            Assert.AreEqual(0, admission.Failures.Count);
            Assert.AreEqual(3, admission.ApprovedActionIds.Count);
        }

        [TestMethod]
        public void AdmissionRejectsExternalTargetSetDriftAfterReview()
        {
            var externalSetId = Guid.NewGuid();
            var reviewed = Plan(
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                externalSetId);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(reviewed, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = TaxonomyAssetApprovalDecision.Approve;
            }
            TaxonomyAssetApprovalFactory.Seal(reviewed, approval, "reviewer@example.com", DateTimeOffset.UtcNow);

            var fresh = TaxonomyAssetContractCloner.Clone(reviewed);
            var replacement = Guid.NewGuid();
            fresh.TermSetProbes.Single().ResolvedTargetTermSetId = replacement;
            fresh.TermProbes.Single().TargetTermSetId = replacement;
            fresh.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(fresh);

            var admission = TaxonomyAssetExecutionAdmissionEvaluator.Evaluate(reviewed, fresh, approval);

            Assert.IsFalse(admission.IsAdmitted);
            Assert.IsTrue(admission.Failures.Any(value => value.Code == "TaxonomyTermSetTargetDrift"));
        }

        [TestMethod]
        public void AdmissionRejectsExternalTermReuseRelationshipDriftAfterReview()
        {
            var externalSetId = Guid.NewGuid();
            var reviewed = Plan(
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                externalSetId);
            var reviewedTerm = reviewed.TermProbes.Single();
            reviewedTerm.ExistingTermSetId = externalSetId;
            reviewedTerm.ExistingTermSetIds = new List<Guid> { externalSetId };
            reviewedTerm.ExistingIsReused = false;
            reviewedTerm.ExistingIsSourceTerm = true;
            reviewed.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(reviewed);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(reviewed, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = TaxonomyAssetApprovalDecision.Approve;
            }
            TaxonomyAssetApprovalFactory.Seal(reviewed, approval, "reviewer@example.com", DateTimeOffset.UtcNow);

            var fresh = TaxonomyAssetContractCloner.Clone(reviewed);
            fresh.TermProbes.Single().ExistingIsReused = true;
            fresh.TermProbes.Single().ExistingIsSourceTerm = false;
            fresh.TermProbes.Single().ExistingReuseSourceTermId = SourceTermId;
            fresh.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(fresh);

            var admission = TaxonomyAssetExecutionAdmissionEvaluator.Evaluate(reviewed, fresh, approval);

            Assert.IsFalse(admission.IsAdmitted);
            Assert.IsTrue(admission.Failures.Any(value => value.Code == "TaxonomyTermRelationshipDrift"));
        }

        [TestMethod]
        public void VerificationRequiresCapturedNativeTermRelationshipRatherThanARepairedReuse()
        {
            var externalSetId = Guid.NewGuid();
            var reviewed = Plan(
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                externalSetId);
            var termPlan = reviewed.Terms.Single();
            termPlan.SourceIsReused = false;
            termPlan.SourceIsSourceTerm = true;
            termPlan.SourceReuseSourceTermId = SourceTermId;
            termPlan.SourceTermSetIds = new List<Guid> { SourceSetId };
            termPlan.SourcePinSourceTermSetId = null;
            termPlan.PlanDigest = TaxonomyAssetIdentity.ComputePlanDigest(termPlan);
            var reviewedProbe = reviewed.TermProbes.Single();
            reviewedProbe.ExistingTermSetId = externalSetId;
            reviewedProbe.ExistingIsReused = false;
            reviewedProbe.ExistingIsSourceTerm = true;
            reviewedProbe.ExistingReuseSourceTermId = SourceTermId;
            reviewedProbe.ExistingTermSetIds = new List<Guid> { externalSetId };
            reviewedProbe.ExistingPinSourceTermSetId = null;
            reviewed.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(reviewed);

            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(reviewed, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = TaxonomyAssetApprovalDecision.Approve;
            }
            TaxonomyAssetApprovalFactory.Seal(reviewed, approval, "reviewer@example.com", DateTimeOffset.UtcNow);

            var matchingFresh = TaxonomyAssetContractCloner.Clone(reviewed);
            var matchingReceipt = SuccessfulReceipt(reviewed, approval);
            TaxonomyAssetVerifier.Verify(reviewed, approval, matchingFresh, matchingReceipt);
            Assert.IsTrue(matchingReceipt.FreshReadbackPassed);

            var repairedFresh = TaxonomyAssetContractCloner.Clone(reviewed);
            repairedFresh.TermProbes.Single().ExistingIsReused = true;
            repairedFresh.TermProbes.Single().ExistingIsSourceTerm = false;
            repairedFresh.TermProbes.Single().ExistingTermSetIds = new List<Guid>
            {
                externalSetId,
                Guid.NewGuid()
            };
            repairedFresh.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(repairedFresh);
            var repairedReceipt = SuccessfulReceipt(reviewed, approval);

            var exception = Assert.ThrowsException<InvalidOperationException>(() =>
                TaxonomyAssetVerifier.Verify(reviewed, approval, repairedFresh, repairedReceipt));

            StringAssert.Contains(exception.Message, "relationship readback differs");
            StringAssert.Contains(exception.Message, "IsReused expected false but observed true");
        }

        [TestMethod]
        public void MaterializationOrdersParentsBeforeChildrenRegardlessOfInputOrder()
        {
            var rootId = Guid.NewGuid();
            var childId = Guid.NewGuid();
            var grandchildId = Guid.NewGuid();
            var root = TermPlan(rootId, null, "root");
            var child = TermPlan(childId, rootId, "root;child");
            var grandchild = TermPlan(grandchildId, childId, "root;child;grandchild");

            var ordered = TaxonomyAssetMaterializationCoordinator
                .OrderTerms(new[] { grandchild, child, root })
                .Select(value => value.Source.TermId)
                .ToArray();

            CollectionAssert.AreEqual(new[] { rootId, childId, grandchildId }, ordered);
        }

        [TestMethod]
        public void OwnedTermSetApprovalCannotHideAnImplicitTermGroupCreation()
        {
            var plan = Plan(
                TaxonomyAssetTargetDisposition.CreateMissing,
                TaxonomyAssetTargetDisposition.CreateMissing,
                SourceSetId,
                TaxonomyAssetTargetDisposition.CreateMissing);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(plan, DateTimeOffset.UtcNow);
            foreach (var action in approval.Actions)
            {
                action.Decision = action.Kind == TaxonomyAssetKind.TermGroup
                    ? TaxonomyAssetApprovalDecision.Defer
                    : TaxonomyAssetApprovalDecision.Approve;
            }

            var exception = Assert.ThrowsException<InvalidDataException>(() => TaxonomyAssetApprovalFactory.Seal(
                plan,
                approval,
                "reviewer@example.com",
                DateTimeOffset.UtcNow));

            StringAssert.Contains(exception.Message, "requires its TermGroup action to be approved");
        }

        [TestMethod]
        public void VerifiedMappingCatalogBindsPlanApprovalReceiptAndPagePlanning()
        {
            var plan = Plan(
                TaxonomyAssetTargetDisposition.ReuseOwned,
                TaxonomyAssetTargetDisposition.ReuseOwned,
                SourceSetId);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(plan, DateTimeOffset.UtcNow);
            TaxonomyAssetApprovalFactory.Seal(plan, approval, "reviewer@example.com", DateTimeOffset.UtcNow);
            var receipt = SuccessfulReceipt(plan, approval);

            var catalog = TaxonomyAssetMappingCatalogFactory.Create(
                plan,
                approval,
                receipt,
                DateTimeOffset.UtcNow);
            var options = new PagePlanningOptions();
            PagePlanningTaxonomyMappingResolver.UseVerifiedCatalog(options, catalog);

            Assert.AreEqual(1, catalog.FieldBindings.Count);
            Assert.AreEqual(receipt.ReceiptDigest, catalog.MaterializationReceiptDigest);
            Assert.AreEqual(SourceSetId, options.TaxonomySchemaMappings.Single().SourceTermSetId);
            Assert.AreEqual(catalog.CatalogDigest, options.TaxonomyAssetMappingCatalog.CatalogDigest);

            options.TaxonomySchemaMappings.Single().TargetTermSetId = Guid.NewGuid();
            Assert.ThrowsException<InvalidDataException>(() =>
                PagePlanningTaxonomyMappingResolver.Normalize(options));
        }

        [TestMethod]
        public void AssessmentMappingsAcceptDeterministicCandidatesButRemoveExternalGuesses()
        {
            var staleGuess = new TaxonomyTargetMapping
            {
                SourceTermStoreId = SourceStoreId,
                SourceTermSetId = SourceSetId,
                TargetTermStoreId = TargetStoreId,
                TargetTermSetId = Guid.NewGuid()
            };

            var reuse = PagePlanningTaxonomyMappingResolver.ResolveForAssessment(
                new[] { staleGuess },
                Plan(
                    TaxonomyAssetTargetDisposition.ReuseOwned,
                    TaxonomyAssetTargetDisposition.ReuseOwned,
                    SourceSetId));
            Assert.AreEqual(1, reuse.Count);
            Assert.AreEqual(SourceSetId, reuse.Single().TargetTermSetId);

            var create = PagePlanningTaxonomyMappingResolver.ResolveForAssessment(
                Array.Empty<TaxonomyTargetMapping>(),
                Plan(
                    TaxonomyAssetTargetDisposition.CreateMissing,
                    TaxonomyAssetTargetDisposition.CreateMissing,
                    SourceSetId));
            Assert.AreEqual(1, create.Count);
            Assert.AreEqual(SourceSetId, create.Single().TargetTermSetId);

            var externalTargetSetId = Guid.NewGuid();
            var external = PagePlanningTaxonomyMappingResolver.ResolveForAssessment(
                new[] { staleGuess },
                Plan(
                    TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                    TaxonomyAssetTargetDisposition.ReviewExternalReuse,
                    externalTargetSetId));
            Assert.AreEqual(0, external.Count);
        }

        [TestMethod]
        public void ReviewPlanDigestValidationIsSafeForConcurrentAssessmentReaders()
        {
            var plan = Plan(
                TaxonomyAssetTargetDisposition.ReuseOwned,
                TaxonomyAssetTargetDisposition.ReuseOwned,
                SourceSetId);
            var expectedReviewDigest = plan.PlanDigest;
            var expectedGroupDigest = plan.TermGroups.Single().PlanDigest;
            var expectedSetDigest = plan.TermSets.Single().PlanDigest;
            var expectedTermDigest = plan.Terms.Single().PlanDigest;
            var failures = new ConcurrentQueue<Exception>();

            Parallel.For(
                0,
                512,
                new ParallelOptions { MaxDegreeOfParallelism = 8 },
                _ =>
                {
                    try
                    {
                        TaxonomyAssetReviewPlanValidator.Validate(plan, true, true);
                        var mappings = PagePlanningTaxonomyMappingResolver.ResolveForAssessment(
                            Array.Empty<TaxonomyTargetMapping>(),
                            plan);
                        if (mappings.Count != 1
                            || mappings.Single().TargetTermSetId != SourceSetId)
                        {
                            throw new InvalidDataException(
                                "Concurrent assessment resolved an unexpected taxonomy mapping.");
                        }
                    }
                    catch (Exception exception)
                    {
                        failures.Enqueue(exception);
                    }
                });

            if (failures.TryPeek(out var failure))
            {
                Assert.Fail(
                    failures.Count + " concurrent taxonomy assessment read(s) failed. First: "
                    + failure.GetType().FullName + ": " + failure.Message);
            }
            Assert.AreEqual(expectedReviewDigest, plan.PlanDigest);
            Assert.AreEqual(expectedGroupDigest, plan.TermGroups.Single().PlanDigest);
            Assert.AreEqual(expectedSetDigest, plan.TermSets.Single().PlanDigest);
            Assert.AreEqual(expectedTermDigest, plan.Terms.Single().PlanDigest);
        }

        [TestMethod]
        public void MaterializationReceiptDigestRejectsIdentityTampering()
        {
            var plan = Plan(
                TaxonomyAssetTargetDisposition.ReuseOwned,
                TaxonomyAssetTargetDisposition.ReuseOwned,
                SourceSetId);
            var approval = TaxonomyAssetApprovalFactory.CreateTemplate(plan, DateTimeOffset.UtcNow);
            TaxonomyAssetApprovalFactory.Seal(plan, approval, "reviewer@example.com", DateTimeOffset.UtcNow);
            var receipt = SuccessfulReceipt(plan, approval);

            receipt.Actions.Single(value => value.Kind == TaxonomyAssetKind.TermSet).TargetTermSetId = Guid.NewGuid();

            Assert.ThrowsException<InvalidDataException>(() =>
                TaxonomyAssetMaterializationReceiptValidator.Validate(plan, approval, receipt));
        }

        private static TaxonomyAssetReviewPlan Plan(
            TaxonomyAssetTargetDisposition setDisposition,
            TaxonomyAssetTargetDisposition termDisposition,
            Guid targetSetId,
            TaxonomyAssetTargetDisposition groupDisposition = TaxonomyAssetTargetDisposition.ReuseOwned)
        {
            var source = new TaxonomyAssetSourceSnapshot
            {
                SourceTenantId = TenantId,
                SnapshotDigest = new string('a', 64),
                TermSets = new List<TaxonomyTermSetSourceSnapshot>
                {
                    new TaxonomyTermSetSourceSnapshot
                    {
                        SourceTenantId = TenantId,
                        SourceTermStoreId = SourceStoreId,
                        SourceTermSetId = SourceSetId,
                        Name = "Wiki Categories",
                        Language = 1033,
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
                        SourceTermStoreId = SourceStoreId,
                        SourceTermSetId = SourceSetId,
                        SourceTermId = SourceTermId,
                        Name = "Proof Points",
                        Path = "Proof Points",
                        Language = 1033,
                        IsAvailableForTagging = true,
                        EvidenceSha256 = new string('c', 64),
                        Availability = EvidenceAvailability.Captured
                    }
                }
            };
            var plan = TaxonomyAssetPlanner.Create(source, TargetStoreId);
            var groupPlan = plan.TermGroups.Single();
            plan.Terms.Single().TargetTermSetId = targetSetId;
            plan.Terms.Single().PlanDigest = TaxonomyAssetIdentity.ComputePlanDigest(plan.Terms.Single());
            plan.TermGroupProbes.Add(new TaxonomyTermGroupTargetProbe
            {
                SourceTenantId = TenantId,
                SourceTermStoreId = SourceStoreId,
                TargetTermStoreId = TargetStoreId,
                ResolvedTargetGroupId = groupDisposition == TaxonomyAssetTargetDisposition.CreateMissing
                    ? null
                    : groupPlan.PreferredTargetGroupId,
                Disposition = groupDisposition
            });
            plan.TermSetProbes.Add(new TaxonomyTermSetTargetProbe
            {
                SourceTermStoreId = SourceStoreId,
                SourceTermSetId = SourceSetId,
                TargetTermStoreId = TargetStoreId,
                ResolvedTargetTermSetId = setDisposition == TaxonomyAssetTargetDisposition.CreateMissing ? null : targetSetId,
                Disposition = setDisposition
            });
            plan.TermProbes.Add(new TaxonomyTermTargetProbe
            {
                SourceTermStoreId = SourceStoreId,
                SourceTermSetId = SourceSetId,
                SourceTermId = SourceTermId,
                TargetTermStoreId = TargetStoreId,
                TargetTermSetId = targetSetId,
                ResolvedTargetTermId = termDisposition == TaxonomyAssetTargetDisposition.ReuseOwned
                    || termDisposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse
                        ? SourceTermId
                        : (Guid?)null,
                Disposition = termDisposition
            });
            plan.MappingCandidates.Add(new TaxonomyAssetMappingCandidate
            {
                SourceTermStoreId = SourceStoreId,
                SourceTermSetId = SourceSetId,
                TargetTermStoreId = TargetStoreId,
                TargetTermSetId = targetSetId,
                Disposition = setDisposition,
                RequiresReview = setDisposition != TaxonomyAssetTargetDisposition.ReuseOwned,
                EvidenceSha256 = new string('d', 64)
            });
            plan.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(plan);
            TaxonomyAssetReviewPlanValidator.Validate(plan);
            return plan;
        }

        private static TaxonomyTermMaterializationPlan TermPlan(Guid termId, Guid? parentId, string path)
        {
            return TaxonomyAssetIdentity.CreateTermPlan(
                new TaxonomyTermSourceSnapshot
                {
                    SourceTenantId = TenantId,
                    SourceTermStoreId = SourceStoreId,
                    SourceTermSetId = SourceSetId,
                    SourceTermId = termId,
                    SourceParentTermId = parentId,
                    Name = path.Split(';').Last(),
                    Path = path,
                    Language = 1033,
                    IsAvailableForTagging = true,
                    EvidenceSha256 = new string('e', 64),
                    Availability = EvidenceAvailability.Captured
                },
                TargetStoreId,
                SourceSetId,
                parentId);
        }

        private static TaxonomyAssetMaterializationReceipt SuccessfulReceipt(
            TaxonomyAssetReviewPlan plan,
            TaxonomyAssetApprovalManifest approval)
        {
            var started = DateTimeOffset.UtcNow.AddSeconds(-1);
            var receipt = new TaxonomyAssetMaterializationReceipt
            {
                OperationId = Guid.NewGuid(),
                ReviewPlanDigest = plan.PlanDigest,
                ApprovalDigest = approval.ApprovalDigest,
                TargetTermStoreId = plan.TargetTermStoreId,
                StartedAtUtc = started,
                CompletedAtUtc = started.AddSeconds(1),
                ChangedTarget = false,
                FreshReadbackPassed = true,
                DeferredActionIds = approval.Actions
                    .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Defer)
                    .Select(value => value.ActionId)
                    .ToList(),
                RejectedActionIds = approval.Actions
                    .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Reject)
                    .Select(value => value.ActionId)
                    .ToList(),
                Actions = approval.Actions
                    .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Approve)
                    .Select(value => new TaxonomyAssetActionReceipt
                    {
                        ActionId = value.ActionId,
                        Kind = value.Kind,
                        SourceTenantId = value.SourceTenantId,
                        SourceTermStoreId = value.SourceTermStoreId,
                        SourceTermSetId = value.SourceTermSetId,
                        SourceTermId = value.SourceTermId,
                        TargetTermStoreId = value.TargetTermStoreId,
                        TargetTermGroupId = value.TargetTermGroupId,
                        TargetTermSetId = value.TargetTermSetId,
                        TargetTermId = value.TargetTermId,
                        ReviewedDisposition = value.ReviewedDisposition,
                        PreflightDisposition = value.ReviewedDisposition,
                        FinalDisposition = TaxonomyAssetVerifier.ExpectedFinalDisposition(value.ReviewedDisposition),
                        ChangedTarget = false,
                        FreshReadbackPassed = true
                    })
                    .ToList()
            };
            TaxonomyAssetMaterializationReceiptValidator.Seal(plan, approval, receipt);
            return receipt;
        }
    }
}
