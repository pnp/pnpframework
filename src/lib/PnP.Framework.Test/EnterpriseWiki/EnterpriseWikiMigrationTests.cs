using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Execution;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.ClassicWebParts.Planning;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.Markup;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Pages.Profiles;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Cohorts;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Packaging;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class EnterpriseWikiMigrationTests
    {
        [TestMethod]
        public void ContentTypeClassificationExcludesProjectPages()
        {
            Assert.IsTrue(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType(BuiltInContentTypeId.EnterpriseWikiPage + "001122"));
            Assert.IsFalse(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType(BuiltInContentTypeId.ProjectPage + "001122"));
            Assert.IsFalse(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType("0x010100C568DB52D9"));
        }

        [TestMethod]
        public void RuntimeResolutionUsesClrPageTypeBeforeContentTypeProfile()
        {
            var artifact = new PageArtifactSnapshot
            {
                PageDirective = new PageDirectiveSnapshot
                {
                    Inherits = "Microsoft.SharePoint.WebPartPages.WikiEditPage, Microsoft.SharePoint"
                }
            };

            var runtime = PageRuntimeResolver.Resolve(
                artifact,
                null,
                BuiltInContentTypeId.EnterpriseWikiPage);

            Assert.AreEqual(PageRuntimeAdapterIds.Wiki, runtime.AdapterId);
            Assert.AreEqual(PageRuntimeDetectionSource.PageDirective, runtime.DetectionSource);
            Assert.AreEqual(PageRuntimeResolutionState.Resolved, runtime.ResolutionState);
        }

        [TestMethod]
        public void ProjectPageProducesNonExclusiveProfileSignalsButIsOutsideEwV1Cohort()
        {
            var source = new PageIdentity
            {
                ContentTypeId = BuiltInContentTypeId.ProjectPage + "001122"
            };
            var signals = PublishingPageProfileSignalProjector.Project(
                source,
                new PublishingPageLayoutSnapshot { FileName = "ProjectPage.aspx" },
                new[]
                {
                    new PageFieldValueSnapshot { InternalName = "Wiki_x0020_Page_x0020_Categories" },
                    new PageFieldValueSnapshot { InternalName = "TaskStatus" }
                });

            Assert.IsTrue(signals.Any(value => value.ProfileId == PageProfileIds.EnterpriseWiki));
            Assert.IsTrue(signals.Any(value => value.ProfileId == PageProfileIds.ProjectPage));
            Assert.AreEqual(
                ValidationCohortDisposition.Excluded,
                EnterpriseWikiV1CohortPolicy.Assess(source.ContentTypeId).Disposition);
        }

        [TestMethod]
        public void RequiredIngredientCanBeDroppedOnlyWhenConsumerTransformReleasesIt()
        {
            var graph = new CanonicalPageIngredientGraph
            {
                Nodes = new List<PageIngredientNode>
                {
                    new PageIngredientNode { Id = "consumer", Kind = PageIngredientKind.Content, HasContent = true },
                    new PageIngredientNode { Id = "dependency", Kind = PageIngredientKind.Reference, HasContent = true }
                },
                Edges = new List<PageIngredientEdge>
                {
                    new PageIngredientEdge
                    {
                        FromIngredientId = "consumer",
                        ToIngredientId = "dependency",
                        Relationship = PageIngredientRelationship.DependsOn,
                        Requirement = PageIngredientRequirement.Required
                    }
                }
            };
            var actions = new List<PageIngredientAction>
            {
                new PageIngredientAction
                {
                    ActionId = "action:consumer",
                    IngredientId = "consumer",
                    Capability = IngredientCapability.Available,
                    Disposition = IngredientDisposition.Preserve
                },
                new PageIngredientAction
                {
                    ActionId = "action:dependency",
                    IngredientId = "dependency",
                    Capability = IngredientCapability.Available,
                    Disposition = IngredientDisposition.Drop
                }
            };

            var blocked = PageIngredientPlanEvaluator.Evaluate(graph, actions);
            Assert.AreEqual(PageMigrationOutcome.Blocked, blocked.Outcome);
            Assert.IsTrue(blocked.Issues.Any(value => value.Code == "RequiredIngredientDependencyUnsatisfied"));

            actions[0].ReleasedDependencyIngredientIds.Add("dependency");
            var invalidRelease = PageIngredientPlanEvaluator.Evaluate(graph, actions);
            Assert.AreEqual(PageMigrationOutcome.Blocked, invalidRelease.Outcome);
            Assert.IsTrue(invalidRelease.Issues.Any(value => value.Code == "IngredientDependencyReleaseInvalid"));

            actions[0].Disposition = IngredientDisposition.Transform;
            var released = PageIngredientPlanEvaluator.Evaluate(graph, actions);
            Assert.AreEqual(PageMigrationOutcome.ExecutableWithLoss, released.Outcome);
            Assert.AreEqual(0, released.Issues.Count);
        }

        [TestMethod]
        public void EmptyIngredientDispositionDoesNotDegradeAggregateOutcome()
        {
            var graph = new CanonicalPageIngredientGraph
            {
                Nodes = new List<PageIngredientNode>
                {
                    new PageIngredientNode { Id = "present", Kind = PageIngredientKind.Content, HasContent = true },
                    new PageIngredientNode { Id = "empty", Kind = PageIngredientKind.Field, HasContent = false }
                }
            };
            var actions = new List<PageIngredientAction>
            {
                new PageIngredientAction
                {
                    ActionId = "action:present",
                    IngredientId = "present",
                    Capability = IngredientCapability.Available,
                    Disposition = IngredientDisposition.Preserve
                },
                new PageIngredientAction
                {
                    ActionId = "action:empty",
                    IngredientId = "empty",
                    Capability = IngredientCapability.Unknown,
                    Disposition = IngredientDisposition.Drop
                }
            };

            Assert.AreEqual(PageMigrationOutcome.Exact, PageIngredientPlanEvaluator.Evaluate(graph, actions).Outcome);
        }

        [TestMethod]
        public void EmptyUnavailableFieldProjectsAsDropInsteadOfBlock()
        {
            var package = CreateMigrationPackage();
            var field = package.Snapshot.Fields.Single(value => value.InternalName == "OOCLReference");
            field.HasValue = false;
            field.Kind = PageFieldValueKind.Null;
            field.RawType = null;
            field.RawValue = null;
            field.RawValueJson = null;
            field.CaptureStatus = PageCaptureStatus.NotReturned;
            package.Snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(package.Snapshot);
            package.Plan.FieldActions.Clear();
            package.Plan.FieldActions.Add(new PageFieldAction
            {
                SourceInternalName = field.InternalName,
                TargetInternalName = field.InternalName,
                Disposition = PageFieldDisposition.CaptureUnavailable,
                Reason = "The field definition was captured, but no source value was returned."
            });

            var actions = PublishingPageIngredientActionProjector.Project(package.Snapshot, package.Plan);
            var fieldAction = actions.Single(value => value.IngredientId == "field:OOCLReference");
            var evaluation = PageIngredientPlanEvaluator.Evaluate(package.Snapshot.IngredientGraph, actions);

            Assert.AreEqual(IngredientCapability.Unknown, fieldAction.Capability);
            Assert.AreEqual(IngredientDisposition.Drop, fieldAction.Disposition);
            Assert.AreEqual("discard-no-source-value", fieldAction.Realization);
            Assert.IsFalse(evaluation.Issues.Any(value => value.Code == "IngredientBlocked"));
        }

        [TestMethod]
        public void DanglingTaxonomyRelationshipIsSealedToExactFieldValuesAndHiddenIdentity()
        {
            var field = CreateDanglingTaxonomyField();

            Assert.AreEqual(0, PageTaxonomyRelationshipEvidence.ValidateSealedField(field).Count);
            Assert.AreEqual(0, PageTaxonomyRelationshipEvidence.GetFidelityErrors(field, field.TaxonomyValues.Single()).Count);
            Assert.AreEqual(TaxonomyRelationshipState.DanglingTermAbsent, field.TaxonomyValues.Single().Relationship.State);
            Assert.IsFalse(string.IsNullOrWhiteSpace(field.TaxonomyValueSetSha256));
            Assert.IsFalse(string.IsNullOrWhiteSpace(field.TaxonomyValues.Single().Relationship.EvidenceSha256));

            field.TaxonomyValues.Single().Label = "healed-label";

            Assert.IsTrue(PageTaxonomyRelationshipEvidence.ValidateSealedField(field).Any(value =>
                value.Contains("digest", StringComparison.OrdinalIgnoreCase)
                || value.Contains("proof", StringComparison.OrdinalIgnoreCase)));
        }

        [TestMethod]
        public void ConflictedTaxonomyBindingRemainsExportableButNotExecutable()
        {
            var snapshot = CreateSnapshot();
            var field = CreateDanglingTaxonomyField();
            field.TaxonomyBinding.TermStoreId = Guid.Empty;
            field.TaxonomyBinding.BoundTermSetId = Guid.Empty;
            field.TaxonomyBinding.TextFieldId = Guid.Empty;
            var relationship = field.TaxonomyValues.Single().Relationship;
            relationship.State = TaxonomyRelationshipState.Conflict;
            relationship.Diagnostics.Add("The source taxonomy field binding could not be read.");
            PageTaxonomyRelationshipProof.Seal(field);
            snapshot.Fields.Add(field);
            snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(snapshot.Source, snapshot.Layout, snapshot.Fields);
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            PublishingPagePackageValidator.ValidateExport(export);

            Assert.AreEqual(0, PageTaxonomyRelationshipEvidence.ValidateSealedField(field).Count);
            Assert.IsTrue(PageTaxonomyRelationshipEvidence.GetFidelityErrors(
                field,
                field.TaxonomyValues.Single()).Any(value =>
                    value.Contains("incomplete", StringComparison.OrdinalIgnoreCase)
                    || value.Contains("Conflict", StringComparison.OrdinalIgnoreCase)));
        }

        [TestMethod]
        public void TaxonomyHiddenListLocalizationPreservesCapturedLcidValues()
        {
            var entry = new TaxonomyHiddenListEntrySnapshot
            {
                Title = "Category",
                Terms = new List<TaxonomyLocalizedTextSnapshot>
                {
                    new TaxonomyLocalizedTextSnapshot { FieldInternalName = "Term1033", Value = "Category" },
                    new TaxonomyLocalizedTextSnapshot { FieldInternalName = "Term2052", Value = "Category ZH" }
                },
                Paths = new List<TaxonomyLocalizedTextSnapshot>
                {
                    new TaxonomyLocalizedTextSnapshot { FieldInternalName = "Path1033", Value = "Root;Category" },
                    new TaxonomyLocalizedTextSnapshot { FieldInternalName = "Path2052", Value = "Root ZH;Category ZH" }
                }
            };
            var targetValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Term1033"] = "Category",
                ["Term2052"] = "Category ZH",
                ["Path1033"] = "Root;Category",
                ["Path2052"] = "Root ZH;Category ZH"
            };

            Assert.AreEqual("Root ZH;Category ZH", entry.PreferredPath("Category ZH"));
            Assert.AreEqual(0, PageTaxonomyHiddenListLocalization.GetTargetCoverageErrors(
                entry,
                new[] { "Term1033", "Term2052" },
                new[] { "Path1033", "Path2052" }).Count);
            Assert.IsTrue(PageTaxonomyHiddenListLocalization.MatchesCapturedValues(
                entry,
                new[] { "Term1033", "Term2052" },
                new[] { "Path1033", "Path2052" },
                fieldName => targetValues[fieldName]));
            Assert.IsTrue(PageTaxonomyHiddenListLocalization.GetTargetCoverageErrors(
                entry,
                new[] { "Term1033", "Term2052" },
                new[] { "Path1033" }).Any(value => value.Contains("Path2052", StringComparison.Ordinal)));
        }

        [TestMethod]
        public void TaxonomyRelationshipIsARequiredIngredientOfThePageField()
        {
            var snapshot = CreateSnapshot();
            var field = CreateDanglingTaxonomyField();
            snapshot.Fields.Add(field);
            snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(snapshot.Source, snapshot.Layout, snapshot.Fields);
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            var termId = Guid.Parse(field.TaxonomyValues.Single().TermGuid);
            var relationshipId = PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, field.TaxonomyValues.Single().WssId);

            var node = snapshot.IngredientGraph.Nodes.Single(value => value.Id == relationshipId);
            var edge = snapshot.IngredientGraph.Edges.Single(value =>
                value.FromIngredientId == PublishingPageIngredientIds.Field(field.InternalName)
                && value.ToIngredientId == relationshipId);

            Assert.AreEqual(PageIngredientKind.Taxonomy, node.Kind);
            Assert.AreEqual(field.TaxonomyValues.Single().Relationship.EvidenceSha256, node.EvidenceDigest);
            Assert.AreEqual(PageIngredientRequirement.Required, edge.Requirement);
        }

        [TestMethod]
        public void ExportValidationRejectsRedigestedTaxonomyRelationshipTampering()
        {
            var snapshot = CreateSnapshot();
            var field = CreateDanglingTaxonomyField();
            snapshot.Fields.Add(field);
            snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(snapshot.Source, snapshot.Layout, snapshot.Fields);
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };
            PublishingPagePackageValidator.ValidateExport(export);

            field.TaxonomyValues.Single().Relationship.ValueHiddenListEntry.Paths.Single().Value = "different/path";
            export.SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void InvalidTaxonomyRelationshipProjectsAsExplicitTransformationWithoutRepair()
        {
            var package = CreateMigrationPackage();
            var field = CreateDanglingTaxonomyField();
            var value = field.TaxonomyValues.Single();
            var termId = Guid.Parse(value.TermGuid);
            package.Snapshot.Fields.Add(field);
            package.Snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(
                package.Snapshot.Source,
                package.Snapshot.Layout,
                package.Snapshot.Fields);
            package.Snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(package.Snapshot);
            package.Plan.FieldActions.Add(new PageFieldAction
            {
                SourceInternalName = field.InternalName,
                TargetInternalName = field.InternalName,
                TargetTypeAsString = field.TypeAsString,
                Disposition = PageFieldDisposition.ApplyTaxonomyRelationships,
                Reason = "Reproduce exact relationship."
            });
            package.Plan.TaxonomyRelationshipActions.Add(new TaxonomyRelationshipAction
            {
                SourceFieldId = field.Id,
                SourceFieldInternalName = field.InternalName,
                SourceTermId = termId,
                SourceWssId = value.WssId,
                SourceEvidenceSha256 = value.Relationship.EvidenceSha256,
                SourceState = TaxonomyRelationshipState.DanglingTermAbsent,
                Disposition = TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent,
                TargetFieldId = field.Id,
                TargetTextFieldId = Guid.Parse("dddddddd-1111-2222-3333-eeeeeeeeeeee"),
                TargetFieldOpen = field.TaxonomyBinding.Open,
                TargetTermStoreId = Guid.Parse("aaaaaaaa-1111-2222-3333-bbbbbbbbbbbb"),
                TargetBoundTermSetId = Guid.Parse("cccccccc-1111-2222-3333-dddddddddddd"),
                TargetValueHiddenListTermSetId = Guid.Parse("cccccccc-1111-2222-3333-dddddddddddd"),
                TargetTaxCatchAllHiddenListTermSetId = Guid.Parse("cccccccc-1111-2222-3333-dddddddddddd"),
                Reason = "Keep the Term absent and reproduce the dangling relationship.",
                VerificationAssertions = new List<string>
                {
                    "The Term remains absent.",
                    "The target-local WssId resolves to the sealed hidden identity."
                }
            });

            var actions = PublishingPageIngredientActionProjector.Project(package.Snapshot, package.Plan);
            var relationship = actions.Single(action => action.IngredientId ==
                PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId));

            Assert.AreEqual(IngredientDisposition.Transform, relationship.Disposition);
            Assert.AreEqual("reproduce-dangling-term-with-target-local-wssid", relationship.Realization);
            StringAssert.Contains(relationship.Reason, "absent");
        }

        [TestMethod]
        public void UnselectedTaxonomyRelationshipRetainsSealedEvidenceWithoutTargetClaims()
        {
            var package = CreateMigrationPackage();
            var field = CreateDanglingTaxonomyField();
            var value = field.TaxonomyValues.Single();
            var termId = Guid.Parse(value.TermGuid);
            package.Snapshot.Fields.Add(field);
            package.Snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(
                package.Snapshot.Source,
                package.Snapshot.Layout,
                package.Snapshot.Fields);
            package.Snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(package.Snapshot);
            package.SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(package.Snapshot);
            package.Plan.SourceSnapshotDigest = package.SnapshotDigest;
            package.Plan.FieldActions.Add(new PageFieldAction
            {
                SourceInternalName = field.InternalName,
                TargetInternalName = field.InternalName,
                Disposition = PageFieldDisposition.EvidenceOnly,
                Reason = "The importer does not own this field."
            });
            package.Plan.TaxonomyRelationshipActions.Add(new TaxonomyRelationshipAction
            {
                SourceFieldId = field.Id,
                SourceFieldInternalName = field.InternalName,
                SourceTermId = termId,
                SourceWssId = value.WssId,
                SourceEvidenceSha256 = value.Relationship.EvidenceSha256,
                SourceState = value.Relationship.State,
                Disposition = TaxonomyRelationshipDisposition.RetainEvidenceOnly,
                Reason = "The owning page field is not selected for replay; its exact taxonomy relationship evidence remains sealed in the snapshot."
            });
            package.Plan.IngredientActions = PublishingPageIngredientActionProjector.Project(package.Snapshot, package.Plan);
            var evaluation = PageIngredientPlanEvaluator.Evaluate(
                package.Snapshot.IngredientGraph,
                package.Plan.IngredientActions);
            package.Plan.MigrationOutcome = evaluation.Outcome;
            package.Plan.IngredientIssues = evaluation.Issues;
            package.State = PublishingPagePackageState.ApprovalReady;
            package.PlanDigest = PublishingPageDigest.ComputePlanDigest(package.Plan);

            PublishingPagePackageValidator.ValidateMigration(package);
            var relationship = package.Plan.IngredientActions.Single(action => action.IngredientId ==
                PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId));

            Assert.AreEqual(IngredientCapability.Unknown, relationship.Capability);
            Assert.AreEqual(IngredientDisposition.Delegate, relationship.Disposition);
            Assert.AreEqual("retain-sealed-relationship-evidence", relationship.Realization);
            Assert.IsNull(relationship.TargetIdentity);
            Assert.AreEqual(PageMigrationOutcome.ExecutableWithLoss, package.Plan.MigrationOutcome);

            package.Plan.TaxonomyRelationshipActions.Single().TargetTermStoreId =
                Guid.Parse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
            package.PlanDigest = PublishingPageDigest.ComputePlanDigest(package.Plan);

            Assert.ThrowsException<InvalidDataException>(() =>
                PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void MigrationPackageRequiresOneSealedActionPerTaxonomyRelationship()
        {
            var package = CreateMigrationPackage();
            var field = CreateDanglingTaxonomyField();
            AddDanglingTaxonomyPlan(package, field);

            PublishingPagePackageValidator.ValidateMigration(package);
            var report = PublishingPageMigrationReportBuilder.Build(package);
            StringAssert.Contains(report, "DanglingTermAbsent");
            StringAssert.Contains(report, "PreserveDanglingTermAbsent");
            StringAssert.Contains(report, "evidenceSha256");

            package.Plan.TaxonomyRelationshipActions.Clear();
            package.PlanDigest = PublishingPageDigest.ComputePlanDigest(package.Plan);

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void ContentRewriteUsesLongestCaseInsensitiveMappingFirst()
        {
            var replacements = new[]
            {
                new PageTextReplacement
                {
                    Source = "https://source.sharepoint.com/sites/source",
                    Target = "https://target.sharepoint.com/sites/target"
                },
                new PageTextReplacement
                {
                    Source = "https://source.sharepoint.com",
                    Target = "https://target.sharepoint.com"
                }
            };

            var actual = PageTextTransformer.Rewrite(
                "<a href=\"HTTPS://SOURCE.SHAREPOINT.COM/sites/source/Pages/A.aspx\">A</a>",
                replacements);

            Assert.AreEqual("<a href=\"https://target.sharepoint.com/sites/target/Pages/A.aspx\">A</a>", actual);
        }

        [TestMethod]
        public void ExportValidationDetectsSnapshotMutation()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            PublishingPagePackageValidator.ValidateExport(export);
            export.Snapshot.PublishingPageContent = "<p>changed</p>";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void ExportValidationRederivesRuntimeEvenAfterSnapshotIsRedigested()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot
            };
            snapshot.Runtime.AdapterId = PageRuntimeAdapterIds.Wiki;
            snapshot.Runtime.DetectionSource = PageRuntimeDetectionSource.PageDirective;
            snapshot.Runtime.ResolutionState = PageRuntimeResolutionState.Resolved;
            export.SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void ExportValidationRederivesIngredientGraphEvenAfterSnapshotIsRedigested()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot
            };
            snapshot.IngredientGraph.Nodes.Single(value => value.Id == PublishingPageIngredientIds.Lifecycle).Label = "tampered";
            export.SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void ExportValidationDetectsPageLayoutByteMutation()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            export.Snapshot.Layout.ContentBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes("changed"));

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void PackageValidationDetectsWorkflowSelectionMutation()
        {
            var package = CreateMigrationPackage();
            package.Selection.ValidationCohort.Disposition = ValidationCohortDisposition.Excluded;

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void ImportPolicyRecomputesCohortAssessmentFromSourceEvidence()
        {
            var package = CreateMigrationPackage();
            package.Selection.ValidationCohort.Reasons = new List<string> { "tampered but re-digested reason" };
            package.SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(package.Selection);

            PublishingPagePackageValidator.ValidateMigration(package);
            Assert.ThrowsException<InvalidDataException>(() =>
                PublishingPageImportPlanValidator.Validate(package, EnterpriseWikiV1WorkflowPolicy.Instance));
        }

        [TestMethod]
        public void MigrationValidationDetectsPlanMutation()
        {
            var package = CreateMigrationPackage();

            PublishingPagePackageValidator.ValidateMigration(package);
            package.Plan.TargetPageServerRelativeUrl = "/sites/target/Pages/changed.aspx";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void MigrationValidationRecomputesIngredientOutcomeAndDependencyIssues()
        {
            var package = CreateMigrationPackage();
            var fieldAction = package.Plan.IngredientActions.Single(value => value.IngredientId == "field:OOCLReference");
            fieldAction.Disposition = IngredientDisposition.Preserve;
            fieldAction.Capability = IngredientCapability.Available;
            package.PlanDigest = PublishingPageDigest.ComputePlanDigest(package.Plan);

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void LifecycleRuleMapsRealR11DraftAndE05PublishedEvidence()
        {
            var r11 = new PageLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Draft",
                ModerationStatus = 3
            };
            var e05 = new PageLifecycleSnapshot
            {
                CheckOutType = "None",
                Level = "Published",
                ModerationStatus = 0
            };

            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(r11));
            Assert.AreEqual(PublishingPageTargetLifecycle.Published, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(e05));
            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(new PageLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Published",
                ModerationStatus = 0
            }));
            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(null));
        }

        [TestMethod]
        public void ReportIncludesEveryCapturedFieldAndItsPlanDisposition()
        {
            var package = CreateMigrationPackage();

            var report = PublishingPageMigrationReportBuilder.Build(package);

            StringAssert.Contains(report, "OOCLReference");
            StringAssert.Contains(report, "Custom recovery field");
            StringAssert.Contains(report, "EvidenceOnly");
            StringAssert.Contains(report, "rawValueJson");
            StringAssert.Contains(report, "Only an unconflicted Published");
            StringAssert.Contains(report, "selection.validationCohort.disposition");
            StringAssert.Contains(report, "CLR runtime resolution");
            StringAssert.Contains(report, "Canonical ingredient nodes");
            StringAssert.Contains(report, "Ingredient actions");
            StringAssert.Contains(report, "snapshot.layout.customizedPageStatus");
            StringAssert.Contains(report, "Page Layout materialization plan");
            StringAssert.Contains(report, "Page Layout target admission");
            StringAssert.Contains(report, "Site collection and Web topology");
            StringAssert.Contains(report, "List dependency closure");
            StringAssert.Contains(report, "Web Part plan actions");
        }

        [TestMethod]
        public void WebPartPortabilityDelegatesListBindingsAndBlocksKnownUnsupportedTypes()
        {
            const string listView = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart"" /></metaData><data><properties><property name=""ListId"">58a84d5d-b1ee-4da0-a49b-7e597ee8ae35</property></properties></data></webPart></webParts>";
            const string rss = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.Portal.WebControls.RSSAggregatorWebPart"" /></metaData></webPart></webParts>";
            const string scriptEditor = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"" /></metaData></webPart></webParts>";

            Assert.IsNull(ClassicWebPartReplayCapabilityPolicy.GetBlocker(listView));
            StringAssert.Contains(ClassicWebPartReplayCapabilityPolicy.GetBlocker(rss), "not supported");
            Assert.IsNull(ClassicWebPartReplayCapabilityPolicy.GetBlocker(scriptEditor));
        }

        [TestMethod]
        public void SerializerRoundTripsTheGenericPublishingPageContract()
        {
            var package = CreateMigrationPackage();

            var json = PublishingPagePackageSerializer.Serialize(package);
            var roundTripped = PublishingPagePackageSerializer.Deserialize<PublishingPageMigrationPackage>(json);

            PublishingPagePackageValidator.ValidateMigration(roundTripped);
            Assert.AreEqual(EnterpriseWikiV1CohortPolicy.CohortId, roundTripped.Selection.WorkflowId);
            Assert.AreEqual(PageRuntimeAdapterIds.Publishing, roundTripped.Snapshot.Runtime.AdapterId);
            Assert.AreEqual(PageMigrationOutcome.ExecutableWithLoss, roundTripped.Plan.MigrationOutcome);
            Assert.AreEqual("https://source.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx", roundTripped.Snapshot.Layout.Url);
            Assert.AreEqual(package.Snapshot.Layout.Bytes.Sha256, roundTripped.Snapshot.Layout.Bytes.Sha256);
            Assert.AreEqual("PublishingPageContent", roundTripped.Snapshot.Layout.Controls.Single().FieldName);
            Assert.AreEqual(package.PlanDigest, roundTripped.PlanDigest);
            Assert.AreEqual(1, roundTripped.Plan.RuntimeVerification.Requirements.Count);
            Assert.AreEqual(RuntimeVerificationRequirementKind.AuthoredDomEquality, roundTripped.Plan.RuntimeVerification.Requirements[0].Kind);
        }

        [TestMethod]
        public void RuntimeVerificationRequirementsAreSealedByThePlanDigest()
        {
            var package = CreateMigrationPackage();

            package.Plan.RuntimeVerification.Requirements[0].Description = "changed after approval";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void ImportReturnsAZeroMutationReceiptWhenThePlanDigestWasNotApproved()
        {
            var package = CreateMigrationPackage();
            var journal = new InMemoryMigrationExecutionJournal();
            using (var context = new ClientContext(package.Plan.TargetWebUrl))
            {
                var receipt = new EnterpriseWikiMigrationImporter().Import(context, package, "not-approved", journal);

                Assert.AreEqual(MigrationExecutionStatus.NotStarted, receipt.ExecutionStatus);
                Assert.AreEqual("PlanDigestNotApproved", receipt.AdmissionFailure.Code);
                Assert.IsFalse(receipt.MutationStarted);
                Assert.AreEqual(StorageVerificationStatus.NotRun, receipt.StorageVerificationStatus);
                Assert.AreEqual(RuntimeVerificationStatus.NotRun, receipt.RuntimeVerificationStatus);
                Assert.AreEqual(MigrationAcceptanceStatus.Rejected, receipt.AcceptanceStatus);
                Assert.AreEqual(0, journal.Intents.Count);
                Assert.AreEqual(1, journal.ExecutionStates.Count);
            }
        }

        [TestMethod]
        public void LayoutMarkupParserCapturesFieldsZonesRegistrationsAndEncodedResources()
        {
            const string markup = @"<%@ Register TagPrefix=""PublishingWebControls"" Namespace=""Microsoft.SharePoint.Publishing.WebControls"" Assembly=""Microsoft.SharePoint.Publishing"" %>
<PublishingWebControls:RichHtmlField ID=""PageContent"" FieldName=""PublishingPageContent"" runat=""server"" />
<WebPartPages:WebPartZone ID=""Main"" runat=""server"" />
<SharePoint:CssRegistration Name=""<% $SPUrl:~sitecollection/Style Library/Contoso/site.css %>"" runat=""server"" />
&lt;script src=&quot;~site/SiteAssets/Contoso/app.js&quot;&gt;&lt;/script&gt;";

            var parsed = PublishingPageLayoutMarkupParser.Parse(markup);

            Assert.AreEqual(1, parsed.Registrations.Count);
            Assert.IsTrue(parsed.RequiredFieldNames.Contains("Title", StringComparer.OrdinalIgnoreCase));
            Assert.IsTrue(parsed.RequiredFieldNames.Contains("PublishingPageContent", StringComparer.OrdinalIgnoreCase));
            Assert.AreEqual("Main", parsed.Zones.Single().Id);
            Assert.IsTrue(parsed.ResourceReferences.Any(value => value.Value == "~sitecollection/Style Library/Contoso/site.css"));
            Assert.IsTrue(parsed.ResourceReferences.Any(value => value.Value == "~site/SiteAssets/Contoso/app.js"));
        }

        [TestMethod]
        public void FieldSchemaCanonicalizerIgnoresStorageSlotsAndRebindsTaxonomy()
        {
            const string left = @"<Field ID=""{11111111-1111-1111-1111-111111111111}"" Name=""Category"" Type=""TaxonomyFieldType"" SourceID=""source-a"" ColName=""nvarchar1"" RowOrdinal=""1""><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value>{22222222-2222-2222-2222-222222222222}</Value></Property><Property><Name>TermSetId</Name><Value>33333333-3333-3333-3333-333333333333</Value></Property><Property><Name>TextField</Name><Value>44444444-4444-4444-4444-444444444444</Value></Property></ArrayOfProperty></Customization></Field>";
            const string right = @"<Field RowOrdinal=""99"" ColName=""nvarchar42"" SourceID=""source-b"" Type=""TaxonomyFieldType"" Name=""Category"" ID=""{11111111-1111-1111-1111-111111111111}""><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value>22222222-2222-2222-2222-222222222222</Value></Property><Property><Name>TermSetId</Name><Value>{33333333-3333-3333-3333-333333333333}</Value></Property><Property><Name>TextField</Name><Value>{44444444-4444-4444-4444-444444444444}</Value></Property></ArrayOfProperty></Customization></Field>";

            Assert.AreEqual(FieldSchemaCanonicalizer.PortableDigest(left), FieldSchemaCanonicalizer.PortableDigest(right));

            var rewritten = FieldSchemaCanonicalizer.RewriteForTarget(
                left,
                Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
                Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"),
                Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc"));

            Assert.IsFalse(rewritten.Contains("ColName"));
            Assert.IsFalse(rewritten.Contains("RowOrdinal"));
            StringAssert.Contains(rewritten, "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            StringAssert.Contains(rewritten, "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            StringAssert.Contains(rewritten, "cccccccc-cccc-cccc-cccc-cccccccccccc");
        }

        [TestMethod]
        public void ContentTypeSchemaPlannerCreatesScalarClosureAndBlocksUnmappedTaxonomy()
        {
            var scalarId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var taxonomyId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var schema = new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                ContentTypeId = "0x010100AA",
                Name = "Custom Page",
                ParentContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC",
                ParentContentTypeName = "Page",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = scalarId, Name = "Activity", Role = FieldSchemaRole.DirectBinding },
                    new ContentTypeFieldLinkSnapshot { FieldId = taxonomyId, Name = "Category", Role = FieldSchemaRole.DirectBinding }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot>
                {
                    Field(scalarId, "Activity", "Text", "<Field ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Activity\" Type=\"Text\" ColName=\"nvarchar1\" />"),
                    new FieldSchemaSnapshot
                    {
                        Id = taxonomyId,
                        InternalName = "Category",
                        Title = "Category",
                        TypeAsString = "TaxonomyFieldType",
                        SchemaXml = "<Field ID=\"{22222222-2222-2222-2222-222222222222}\" Name=\"Category\" Type=\"TaxonomyFieldType\" />",
                        PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest("<Field ID=\"{22222222-2222-2222-2222-222222222222}\" Name=\"Category\" Type=\"TaxonomyFieldType\" />"),
                        Role = FieldSchemaRole.DirectBinding,
                        Taxonomy = new TaxonomyFieldBindingSnapshot
                        {
                            SourceTermStoreId = Guid.Parse("33333333-3333-3333-3333-333333333333"),
                            SourceTermSetId = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                            HiddenTextFieldId = Guid.Parse("55555555-5555-5555-5555-555555555555")
                        }
                    }
                }
            };

            var plan = ContentTypeSchemaPlanner.CreateRequiredClosure(schema);

            Assert.AreEqual(ContentTypeMaterializationDisposition.Block, plan.Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                plan.Fields.Single(value => value.FieldId == scalarId).Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.Block,
                plan.Fields.Single(value => value.FieldId == taxonomyId).Disposition);
            Assert.IsFalse(plan.Fields.Single(value => value.FieldId == scalarId).TargetSchemaXml.Contains("ColName"));
        }

        [TestMethod]
        public void CustomLayoutPlanMapsExactSchemaResourcesAndDeterministicTargetBytes()
        {
            var layout = CreateCustomLayout();

            var plan = PublishingPageLayoutPlanFactory.Create(
                layout,
                new Uri("https://source.sharepoint.com/sites/source"),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");

            Assert.AreEqual(PublishingPageLayoutMaterializationDisposition.CreateOwned, plan.Disposition);
            StringAssert.StartsWith(plan.TargetFileName, "pnp-custom-");
            StringAssert.EndsWith(plan.TargetFileName, ".aspx");
            Assert.AreEqual("/sites/target/_catalogs/masterpage/" + plan.TargetFileName, plan.TargetServerRelativeUrl);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, plan.ContentTypeSchema.Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.CreateOrReuseOwned, plan.ContentTypeSchema.Fields.Single().Disposition);
            Assert.AreEqual("/sites/target/SiteAssets/Contoso/app.js", plan.ResourceMaterializations.Single().TargetServerRelativeUrl);
            Assert.AreEqual("https://target.sharepoint.com/sites/target/SiteAssets/Contoso/app.js", plan.ResourceRewrites.Single().TargetReference);
            Assert.AreNotEqual(plan.SourceBytes.Sha256, plan.TargetBytes.Sha256);
        }

        [TestMethod]
        public void CustomLayoutAdmissionCreatesOnlyWhenSchemaResourcesAndTargetAreEligible()
        {
            var layout = CreateCustomLayout();
            var plan = PublishingPageLayoutPlanFactory.Create(
                layout,
                new Uri("https://source.sharepoint.com/sites/source"),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");
            var probe = CreateEligibleCustomLayoutProbe(plan);

            var admission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsTrue(admission.IsEligible);
            Assert.AreEqual(PublishingPageLayoutMaterializationDisposition.CreateOwned, admission.Disposition);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, admission.ContentTypeSchema.Disposition);

            probe.Resources.Single().FileExists = true;
            probe.Resources.Single().ExistingBytesSha256 = new string('0', 64);
            var collision = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, probe);
            Assert.IsFalse(collision.IsEligible);
            Assert.IsTrue(collision.Issues.Any(value => value.Code == "TargetLayoutResourceCollision"));
        }

        [TestMethod]
        public void ExportValidationRequiresOneEvidenceRecordPerLayoutResourceReference()
        {
            var snapshot = CreateSnapshot();
            snapshot.Layout = CreateCustomLayout();
            snapshot.Layout.ResourceArtifacts.Clear();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void DirectoryArtifactStoreRoundTripsAndDeduplicatesByDigest()
        {
            var root = Path.Combine(Path.GetTempPath(), "pnp-migration-artifacts-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new DirectoryMigrationArtifactStore(root);
                var bytes = Encoding.UTF8.GetBytes("exact migration payload");
                ArtifactReference first;
                ArtifactReference second;
                using (var content = new MemoryStream(bytes, false))
                {
                    first = store.Put(content, "text/plain", "payload.txt");
                }

                using (var content = new MemoryStream(bytes, false))
                {
                    second = store.Put(content, "text/plain", "another-name.txt");
                }

                Assert.AreEqual(first.Sha256, second.Sha256);
                Assert.AreEqual(bytes.LongLength, first.Length);
                Assert.IsTrue(store.Contains(first.Sha256));
                using (var content = store.OpenRead(first.Sha256))
                using (var buffer = new MemoryStream())
                {
                    content.CopyTo(buffer);
                    CollectionAssert.AreEqual(bytes, buffer.ToArray());
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [TestMethod]
        public void ListWebPartBindingRewritesListWebViewAndPageIdentities()
        {
            var sourceWeb = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var sourceList = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var sourceView = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var targetWeb = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            var targetList = Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            var targetView = Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc");
            var exportXml = "<webParts><webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\"><metaData><type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart\" /></metaData><data><properties>"
                + "<property name=\"ListId\">" + sourceList.ToString("D") + "</property>"
                + "<property name=\"ListName\">{" + sourceList.ToString("D") + "}</property>"
                + "<property name=\"WebId\">00000000-0000-0000-0000-000000000000</property>"
                + "<property name=\"ViewGuid\">" + sourceView.ToString("D") + "</property>"
                + "<property name=\"TitleUrl\">/teams/source/child/Lists/Resources</property>"
                + "<property name=\"XmlDefinition\">&lt;View Name=\"{" + sourceView.ToString("D") + "}\" Url=\"/teams/source/child/Pages/A.aspx\"&gt;&lt;JSLink&gt;clienttemplates.js&lt;/JSLink&gt;&lt;/View&gt;</property>"
                + "</properties></data></webPart></webParts>";
            var snapshot = new ClassicWebPartSnapshot
            {
                Id = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                Title = "Resources",
                TypeName = "Microsoft.SharePoint.WebPartPages.XsltListViewWebPart",
                ExportXml = exportXml,
                ExportSha256 = MigrationDigest.ComputeSha256(exportXml)
            };

            var parsed = ClassicListWebPartBindingParser.Parse(
                snapshot,
                sourceWeb,
                "https://source.sharepoint.com/teams/source/child",
                "/teams/source/child/Pages/A.aspx");

            Assert.IsTrue(parsed.IsExecutable, string.Join(Environment.NewLine, parsed.Issues.Select(value => value.Message)));
            Assert.AreEqual(sourceWeb, parsed.Binding.SourceListWebId);
            Assert.AreEqual(sourceView, parsed.Binding.SourceViewId);
            Assert.AreEqual("clienttemplates.js", parsed.Binding.JsLink);

            var rewritten = ClassicListWebPartRewriter.Rewrite(parsed.Binding, new ClassicListWebPartTargetMap
            {
                SourceWebId = sourceWeb,
                SourceListId = sourceList,
                SourceViewId = sourceView,
                TargetWebId = targetWeb,
                TargetListId = targetList,
                TargetViewId = targetView,
                TargetListServerRelativeUrl = "/sites/target/child/Lists/Resources",
                TargetListAbsoluteUrl = "https://target.sharepoint.com/sites/target/child/Lists/Resources",
                TargetPageServerRelativeUrl = "/sites/target/child/Pages/A.aspx"
            });
            var properties = XDocument.Parse(rewritten.ExportXml).Descendants()
                .Where(value => value.Name.LocalName == "property")
                .ToDictionary(value => (string)value.Attribute("name"), value => value.Value, StringComparer.OrdinalIgnoreCase);

            Assert.AreEqual(targetWeb.ToString("D"), properties["WebId"]);
            Assert.AreEqual(targetList.ToString("D"), properties["ListId"]);
            Assert.AreEqual("{" + targetList.ToString("D") + "}", properties["ListName"]);
            Assert.AreEqual(targetView.ToString("D"), properties["ViewGuid"]);
            var view = XDocument.Parse(properties["XmlDefinition"]).Root;
            Assert.AreEqual("{" + targetView.ToString("D") + "}", (string)view.Attribute("Name"));
            Assert.AreEqual("/sites/target/child/Pages/A.aspx", (string)view.Attribute("Url"));
        }

        [TestMethod]
        public void LookupDependencyGraphOrdersLookupListsAndBlocksCycles()
        {
            var owner = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var lookup = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var ordered = ListLookupDependencyGraph.Order(
                new[] { owner, lookup },
                new[]
                {
                    new ListLookupDependency
                    {
                        SourceListId = owner,
                        LookupListId = lookup,
                        FieldId = Guid.NewGuid(),
                        FieldInternalName = "Lookup"
                    }
                });

            Assert.IsTrue(ordered.IsExecutable);
            CollectionAssert.AreEqual(new[] { lookup, owner }, ordered.OrderedSourceListIds.ToArray());

            var cycle = ListLookupDependencyGraph.Order(
                new[] { owner, lookup },
                new[]
                {
                    new ListLookupDependency { SourceListId = owner, LookupListId = lookup },
                    new ListLookupDependency { SourceListId = lookup, LookupListId = owner }
                });
            Assert.IsFalse(cycle.IsExecutable);
            Assert.IsTrue(cycle.Issues.Any(value => value.Code == "LookupDependencyCycle"));
        }

        [TestMethod]
        public void TopologyPlannerPreservesNestedWebOwnershipAndStableDigest()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var rootId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var childId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var source = new SourceSiteCollectionSnapshot
            {
                SiteId = siteId,
                SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                ServerRelativeUrl = "/teams/source",
                RootWebId = rootId,
                Webs = new List<SourceWebSnapshot>
                {
                    new SourceWebSnapshot
                    {
                        SiteId = siteId,
                        WebId = childId,
                        ParentWebId = rootId,
                        SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                        WebUrl = "https://source.sharepoint.com/teams/source/child",
                        ServerRelativeUrl = "/teams/source/child",
                        Title = "Child",
                        WebTemplate = "CMSPUBLISHING"
                    },
                    new SourceWebSnapshot
                    {
                        SiteId = siteId,
                        WebId = rootId,
                        SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                        WebUrl = "https://source.sharepoint.com/teams/source",
                        ServerRelativeUrl = "/teams/source",
                        Title = "Root",
                        WebTemplate = "CMSPUBLISHING"
                    }
                }
            };
            var target = new TargetSiteCollectionSpec
            {
                SourceSiteId = siteId,
                Mode = TargetSiteMode.ExistingTargetSite,
                TargetSiteUrl = "https://target.sharepoint.com/sites/target",
                ExpectedTargetSiteId = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
                Title = "Target"
            };
            var policy = new TopologyPlanningPolicy
            {
                WebOverrides = new List<TargetWebOverride>
                {
                    new TargetWebOverride { SourceWebId = childId, TargetUrlSegment = "area", TargetTitle = "Area" }
                }
            };

            var first = new TopologyPlanner().Build(new[] { source }, new[] { target }, policy);
            var second = new TopologyPlanner().Build(new[] { source }, new[] { target }, policy);

            Assert.IsTrue(first.IsExecutable, string.Join(Environment.NewLine, first.Issues.Select(value => value.Message)));
            var child = first.Plan.SiteCollections.Single().Webs.Single(value => value.SourceWebId == childId);
            Assert.AreEqual("/sites/target/area", child.TargetServerRelativeUrl);
            Assert.AreEqual("/sites/target/area/Lists/Resources", TopologyPlanner.MapWebOwnedServerRelativePath(
                "/teams/source/child/Lists/Resources",
                "/teams/source/child",
                child.TargetServerRelativeUrl));
            Assert.AreEqual(first.Plan.PlanDigest, second.Plan.PlanDigest);
        }

        [TestMethod]
        public void ListItemValueCaptureKeepsUnsupportedRawEvidenceForFutureRecovery()
        {
            var captured = ListItemValueSerializer.Serialize("FutureField", new Dictionary<string, object>
            {
                ["reference"] = "OOCL-42",
                ["sequence"] = 7
            });

            Assert.AreEqual(ListItemValueKind.Unsupported, captured.Kind);
            Assert.AreEqual(EvidenceAvailability.Partial, captured.Availability);
            Assert.AreEqual(typeof(Dictionary<string, object>).FullName, captured.RawType);
            StringAssert.Contains(captured.RawValueJson, "OOCL-42");
            Assert.IsTrue(captured.Diagnostics.Any(value => value.Contains("No typed list-item serializer")));
        }

        [TestMethod]
        public void ListPlannerOrdersLookupClosureAndRetainsUnusedUnknownFieldsAsEvidence()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var ownerId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var lookupId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var lookupFieldId = Guid.Parse("55555555-5555-5555-5555-555555555555");
            var unknownFieldId = Guid.Parse("66666666-6666-6666-6666-666666666666");
            var owner = CreateListSnapshot(siteId, webId, ownerId, "Owner");
            owner.Fields.Add(new ListFieldSnapshot
            {
                Id = lookupFieldId,
                InternalName = "Category",
                TypeAsString = "Lookup",
                SchemaXml = "<Field ID='{55555555-5555-5555-5555-555555555555}' Name='Category' Type='Lookup' List='{44444444-4444-4444-4444-444444444444}' ShowField='Title' />",
                SourceLookupWebId = webId,
                SourceLookupListId = lookupId,
                LookupField = "Title"
            });
            owner.Fields.Add(new ListFieldSnapshot
            {
                Id = unknownFieldId,
                InternalName = "FutureValue",
                TypeAsString = "FutureType",
                SchemaXml = "<Field ID='{66666666-6666-6666-6666-666666666666}' Name='FutureValue' Type='FutureType' />"
            });
            var lookup = CreateListSnapshot(siteId, webId, lookupId, "Lookup");

            var plan = ListMigrationPlanFactory.Create(
                new[] { owner, lookup },
                new[] { new ListLookupDependency { SourceListId = ownerId, LookupListId = lookupId, FieldId = lookupFieldId, FieldInternalName = "Category" } },
                CreateTopology(siteId, webId),
                null,
                null);

            Assert.IsTrue(plan.IsExecutable, string.Join(Environment.NewLine, plan.Issues.Select(value => value.Message)));
            CollectionAssert.AreEqual(new[] { lookupId, ownerId }, plan.OrderedSourceListIds.ToArray());
            var ownerPlan = plan.Lists.Single(value => value.SourceListId == ownerId);
            Assert.AreEqual(ListFieldMaterializationDisposition.MapLookup, ownerPlan.Fields.Single(value => value.SourceFieldId == lookupFieldId).Disposition);
            Assert.AreEqual(ListFieldMaterializationDisposition.EvidenceOnly, ownerPlan.Fields.Single(value => value.SourceFieldId == unknownFieldId).Disposition);
        }

        [TestMethod]
        public void IngredientProjectionAssignsExplicitActionsAcrossTopologyAndListContents()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var libraryId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var titleFieldId = Guid.Parse("55555555-5555-5555-5555-555555555555");
            var futureFieldId = Guid.Parse("66666666-6666-6666-6666-666666666666");
            var libraryFieldId = Guid.Parse("77777777-7777-7777-7777-777777777777");
            const string siteContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
            const string listContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";
            var attachmentBytes = Encoding.UTF8.GetBytes("attachment");
            var documentBytes = Encoding.UTF8.GetBytes("document");
            var list = CreateListSnapshot(siteId, webId, listId, "Items");
            list.EnableAttachments = true;
            list.Fields.Add(new ListFieldSnapshot
            {
                Id = titleFieldId,
                InternalName = "CaseTitle",
                Title = "Case title",
                TypeAsString = "Text",
                SchemaXml = "<Field ID='{55555555-5555-5555-5555-555555555555}' Name='CaseTitle' Type='Text' />"
            });
            list.Fields.Add(new ListFieldSnapshot
            {
                Id = futureFieldId,
                InternalName = "FutureValue",
                Title = "Future value",
                TypeAsString = "FutureType",
                SchemaXml = "<Field ID='{66666666-6666-6666-6666-666666666666}' Name='FutureValue' Type='FutureType' />"
            });
            list.ContentTypes.Add(new ListContentTypeSnapshot { Id = "0x01", Name = "Item" });
            list.SourceItemCount = 1;
            list.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 7,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "CaseTitle", Kind = ListItemValueKind.String, ScalarValue = "OOCL" }
                },
                Attachments = new List<ListAttachmentSnapshot>
                {
                    new ListAttachmentSnapshot
                    {
                        FileName = "evidence.txt",
                        ServerRelativeUrl = "/sites/source/Lists/Items/Attachments/7/evidence.txt",
                        Content = new ListBinaryArtifactSnapshot
                        {
                            Artifact = MigrationArtifact.Describe(attachmentBytes, "text/plain", "evidence.txt"),
                            ContentBase64 = Convert.ToBase64String(attachmentBytes),
                            Availability = EvidenceAvailability.Captured
                        }
                    }
                }
            });
            var library = CreateListSnapshot(siteId, webId, libraryId, "Documents");
            library.BaseTemplate = 101;
            library.BaseType = "DocumentLibrary";
            library.RootFolderServerRelativeUrl = "/sites/source/Documents";
            var libraryFieldSchema = "<Field ID='{77777777-7777-7777-7777-777777777777}' Name='CaseNumber' Type='Text' />";
            library.Fields.Add(new ListFieldSnapshot
            {
                Id = libraryFieldId,
                InternalName = "CaseNumber",
                Title = "Case number",
                TypeAsString = "Text",
                SchemaXml = libraryFieldSchema,
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(libraryFieldSchema)
            });
            library.SiteContentTypes.Add(new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceScope = "/sites/source",
                ContentTypeId = siteContentTypeId,
                Name = "Case document",
                ParentContentTypeId = "0x0101",
                ParentContentTypeName = "Document",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = libraryFieldId, Name = "CaseNumber" }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot>
                {
                    new FieldSchemaSnapshot
                    {
                        Id = libraryFieldId,
                        InternalName = "CaseNumber",
                        Title = "Case number",
                        TypeAsString = "Text",
                        SchemaXml = libraryFieldSchema,
                        PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(libraryFieldSchema),
                        Role = FieldSchemaRole.DirectBinding
                    }
                }
            });
            library.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = listContentTypeId,
                Name = "Case document",
                ParentId = siteContentTypeId,
                FieldLinks = new List<ListContentTypeFieldLinkSnapshot>
                {
                    new ListContentTypeFieldLinkSnapshot { FieldId = libraryFieldId, InternalName = "CaseNumber" }
                }
            });
            library.SourceItemCount = 1;
            library.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 9,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "CaseNumber", Kind = ListItemValueKind.String, ScalarValue = "C-42" }
                },
                Document = new ListDocumentSnapshot
                {
                    Kind = ListDocumentObjectKind.File,
                    Name = "case.docx",
                    ServerRelativeUrl = "/sites/source/Documents/case.docx",
                    Length = documentBytes.LongLength,
                    Content = new ListBinaryArtifactSnapshot
                    {
                        Artifact = MigrationArtifact.Describe(documentBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "case.docx"),
                        ContentBase64 = Convert.ToBase64String(documentBytes),
                        Availability = EvidenceAvailability.Captured
                    }
                }
            });

            var package = CreateMigrationPackage();
            var snapshot = package.Snapshot;
            snapshot.Source.SiteId = siteId;
            snapshot.Source.WebId = webId;
            snapshot.SourceTopology = new SourceSiteCollectionSnapshot
            {
                SiteId = siteId,
                RootWebId = webId,
                SiteCollectionUrl = snapshot.Source.WebUrl,
                ServerRelativeUrl = snapshot.Source.WebServerRelativeUrl,
                Webs = new List<SourceWebSnapshot>
                {
                    new SourceWebSnapshot
                    {
                        SiteId = siteId,
                        WebId = webId,
                        SiteCollectionUrl = snapshot.Source.WebUrl,
                        WebUrl = snapshot.Source.WebUrl,
                        ServerRelativeUrl = snapshot.Source.WebServerRelativeUrl,
                        Title = "Source",
                        WebTemplate = "STS",
                        Availability = EvidenceAvailability.Captured
                    }
                }
            };
            snapshot.ListDependencies = new List<ListDependencySnapshot> { list, library };
            var topology = CreateTopology(siteId, webId);
            package.Plan.Topology = topology;
            package.Plan.TopologyTargetAnalysis = CreateAdmittedTopologyAnalysis(topology, siteId, webId);
            package.Plan.ListMigration = ListMigrationPlanFactory.Create(snapshot.ListDependencies, null, topology, null, null);
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);

            var actions = PublishingPageIngredientActionProjector.Project(snapshot, package.Plan);
            var byId = actions.ToDictionary(value => value.IngredientId, StringComparer.Ordinal);

            Assert.IsFalse(actions.Any(value => value.PolicyId == "policy.ingredient.unknown"));
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.Web(siteId, webId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListField(webId, listId, titleFieldId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Drop, byId[PublishingPageIngredientIds.ListField(webId, listId, futureFieldId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListContentType(webId, listId, "0x01")].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListItem(webId, listId, 7)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListAttachment(webId, listId, 7, "evidence.txt")].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.SiteContentType("/sites/source", siteContentTypeId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.SiteField("/sites/source", libraryFieldId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListContentType(webId, libraryId, listContentTypeId)].Disposition);
            Assert.AreEqual(IngredientDisposition.Preserve, byId[PublishingPageIngredientIds.ListDocument(webId, libraryId, 9)].Disposition);
            Assert.AreEqual(0, PageIngredientPlanEvaluator.Evaluate(snapshot.IngredientGraph, actions).Issues.Count);
        }

        [TestMethod]
        public void IngredientProjectionCollapsesOnlySemanticallyIdenticalEdges()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var fieldId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var list = CreateListSnapshot(siteId, webId, listId, "Items");
            list.Fields.Add(new ListFieldSnapshot
            {
                Id = fieldId,
                InternalName = "RepeatedValue",
                Title = "Repeated value",
                TypeAsString = "Text",
                SchemaXml = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='RepeatedValue' Type='Text' />"
            });
            list.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 7,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "RepeatedValue", Kind = ListItemValueKind.String, ScalarValue = "first" },
                    new ListItemValueSnapshot { InternalName = "RepeatedValue", Kind = ListItemValueKind.String, ScalarValue = "second" }
                }
            });

            var snapshot = CreateMigrationPackage().Snapshot;
            snapshot.ListDependencies = new List<ListDependencySnapshot> { list };
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);

            var itemId = PublishingPageIngredientIds.ListItem(webId, listId, 7);
            var listFieldId = PublishingPageIngredientIds.ListField(webId, listId, fieldId);
            Assert.AreEqual(1, snapshot.IngredientGraph.Edges.Count(edge =>
                edge.FromIngredientId == itemId
                && edge.ToIngredientId == listFieldId
                && edge.Relationship == PageIngredientRelationship.BindsTo
                && edge.Requirement == PageIngredientRequirement.Required));
        }

        [TestMethod]
        public void ListPackageValidationAllowsSiblingContentTypesWithTheSameParent()
        {
            var list = CreateListSnapshot(
                Guid.Parse("11111111-1111-1111-1111-111111111111"),
                Guid.Parse("22222222-2222-2222-2222-222222222222"),
                Guid.Parse("33333333-3333-3333-3333-333333333333"),
                "Academy");
            list.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                Name = "Academy article",
                ParentId = "0x01"
            });
            list.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = "0x0100BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB",
                Name = "Academy announcement",
                ParentId = "0x01"
            });

            ListDependencyPackageValidator.Validate(
                Array.Empty<ClassicWebPartSnapshot>(),
                Array.Empty<ClassicListWebPartBindingSnapshot>(),
                new[] { list },
                Array.Empty<ListLookupDependency>(),
                null,
                null);
        }

        [TestMethod]
        public void ListPackageValidationRetainsBindingEvidenceWhenTopologyCaptureIsUnavailable()
        {
            var webPartId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var siteId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var webId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var listId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            const string exportXml = "<webPart />";
            const string exportDigest = "captured-export-digest";
            var list = CreateListSnapshot(siteId, webId, listId, "Referenced list");
            var webPart = new ClassicWebPartSnapshot
            {
                Id = webPartId,
                ExportXml = exportXml,
                ExportSha256 = exportDigest
            };
            var binding = new ClassicListWebPartBindingSnapshot
            {
                SourceWebPartId = webPartId,
                SourceListWebId = webId,
                SourceListId = listId,
                SourceExportXml = exportXml,
                SourceExportSha256 = exportDigest
            };

            ListDependencyPackageValidator.Validate(
                new[] { webPart },
                new[] { binding },
                new[] { list },
                Array.Empty<ListLookupDependency>(),
                null,
                null);
        }

        [TestMethod]
        public void ListPackageValidationRetainsExplicitDocumentLengthMismatchEvidence()
        {
            var bytes = Encoding.UTF8.GetBytes("captured payload");
            var list = CreateListSnapshot(
                Guid.Parse("11111111-1111-1111-1111-111111111111"),
                Guid.Parse("22222222-2222-2222-2222-222222222222"),
                Guid.Parse("33333333-3333-3333-3333-333333333333"),
                "Documents");
            list.BaseType = "DocumentLibrary";
            list.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 7,
                Availability = EvidenceAvailability.Partial,
                Document = new ListDocumentSnapshot
                {
                    Kind = ListDocumentObjectKind.File,
                    Name = "changing.pptx",
                    ServerRelativeUrl = "/sites/source/Documents/changing.pptx",
                    Length = bytes.LongLength + 10,
                    Content = new ListBinaryArtifactSnapshot
                    {
                        Artifact = MigrationArtifact.Describe(bytes, "application/octet-stream", "changing.pptx"),
                        ContentBase64 = Convert.ToBase64String(bytes),
                        Availability = EvidenceAvailability.Partial,
                        Diagnostics = { "DocumentMetadataLengthMismatch: metadataLength=26; payloadLength=16." }
                    }
                }
            });

            ListDependencyPackageValidator.Validate(
                Array.Empty<ClassicWebPartSnapshot>(),
                Array.Empty<ClassicListWebPartBindingSnapshot>(),
                new[] { list },
                Array.Empty<ListLookupDependency>(),
                null,
                null);
        }

        [TestMethod]
        public void IngredientProjectionAssignsLayoutResourceActionFromMaterializationPlan()
        {
            var package = CreateMigrationPackage();
            var snapshot = package.Snapshot;
            snapshot.Layout = CreateCustomLayout();
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            package.Plan.LayoutMaterialization = PublishingPageLayoutPlanFactory.Create(
                snapshot.Layout,
                new Uri(snapshot.Source.WebUrl),
                new Uri(package.Plan.TargetWebUrl),
                new Uri(package.Plan.TargetWebUrl),
                "EnterpriseWiki.aspx");
            package.Plan.LayoutAdmission = null;

            var actions = PublishingPageIngredientActionProjector.Project(snapshot, package.Plan);
            var resource = actions.Single(value => value.IngredientId == PublishingPageIngredientIds.LayoutResource("~site/SiteAssets/Contoso/app.js"));

            Assert.AreEqual(IngredientDisposition.Preserve, resource.Disposition);
            Assert.AreEqual("copy-exact-bytes-create-only", resource.Realization);
            Assert.AreEqual("policy.layout.resource", resource.PolicyId);
            Assert.IsFalse(actions.Any(value => value.PolicyId == "policy.ingredient.unknown"));
        }

        [TestMethod]
        public void ListPlannerBlocksNonemptyPrincipalValuesWithoutExplicitMapping()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var fieldId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var source = CreateListSnapshot(siteId, webId, listId, "People");
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = fieldId,
                InternalName = "Owner",
                TypeAsString = "User",
                SchemaXml = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Owner' Type='User' />"
            });
            source.SourceItemCount = 1;
            source.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 1,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "Owner", Kind = ListItemValueKind.User, ScalarValue = "i:0#.f|membership|owner@example.com" }
                }
            });

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);

            Assert.IsFalse(plan.IsExecutable);
            Assert.AreEqual(ListFieldMaterializationDisposition.Block, plan.Lists.Single().Fields.Single(value => value.SourceFieldId == fieldId).Disposition);
            Assert.IsTrue(plan.Lists.Single().Issues.Any(value => value.Code == "PrincipalMappingUnavailable"));
        }

        [TestMethod]
        public void WebPartReplayCompositionUsesMaterializedListAndViewReceipts()
        {
            var sourceWeb = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var sourceList = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var sourceView = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var webPartId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var targetWeb = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            var targetList = Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            var targetView = Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc");
            var xml = "<webParts><webPart><data><properties>"
                + "<property name='ListId'>" + sourceList.ToString("D") + "</property>"
                + "<property name='ListName'>{" + sourceList.ToString("D") + "}</property>"
                + "<property name='WebId'>" + sourceWeb.ToString("D") + "</property>"
                + "<property name='ViewGuid'>" + sourceView.ToString("D") + "</property>"
                + "<property name='TitleUrl'>/sites/source/Lists/Items</property>"
                + "<property name='XmlDefinition'>&lt;View Name=\"{" + sourceView.ToString("D") + "}\" Url=\"/sites/source/Pages/A.aspx\" /&gt;</property>"
                + "</properties></data></webPart></webParts>";
            var captured = new ClassicWebPartSnapshot { Id = webPartId, ExportXml = xml };
            var binding = new ClassicListWebPartBindingSnapshot
            {
                SourceWebPartId = webPartId,
                SourceListWebId = sourceWeb,
                SourceListId = sourceList,
                SourceViewId = sourceView,
                SourceTitleUrl = "/sites/source/Lists/Items",
                SourceExportXml = xml
            };
            var receipt = new ListMaterializationReceipt
            {
                SourceWebId = sourceWeb,
                SourceListId = sourceList,
                TargetWebId = targetWeb,
                TargetListId = targetList,
                TargetRootFolderServerRelativeUrl = "/sites/target/Lists/Items",
                TargetViewIds = new Dictionary<Guid, Guid> { [sourceView] = targetView }
            };

            var replay = ClassicWebPartReplayComposer.Compose(
                captured,
                new ClassicWebPartAction { SourceWebPartId = webPartId, Disposition = ClassicWebPartDisposition.RebindListAfterMaterialization },
                binding,
                receipt,
                "/sites/target/Pages/A.aspx",
                Array.Empty<PageTextReplacement>());
            var properties = XDocument.Parse(replay).Descendants().Where(value => value.Name.LocalName == "property")
                .ToDictionary(value => (string)value.Attribute("name"), value => value.Value, StringComparer.OrdinalIgnoreCase);

            Assert.AreEqual(targetWeb.ToString("D"), properties["WebId"]);
            Assert.AreEqual(targetList.ToString("D"), properties["ListId"]);
            Assert.AreEqual(targetView.ToString("D"), properties["ViewGuid"]);
            Assert.AreEqual("/sites/target/Lists/Items", properties["TitleUrl"]);
        }

        [TestMethod]
        public void CustomDocumentContentTypeIsCapturedAsClosureInsteadOfMisclassifiedAsRuntime()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var fieldId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            const string siteContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
            const string listContentTypeId = siteContentTypeId + "00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";
            var fieldSchema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='CaseNumber' DisplayName='Case number' Type='Text' />";
            var source = CreateListSnapshot(siteId, webId, listId, "Documents");
            source.BaseTemplate = 101;
            source.BaseType = "DocumentLibrary";
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = fieldId,
                InternalName = "CaseNumber",
                TypeAsString = "Text",
                SchemaXml = fieldSchema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(fieldSchema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(fieldSchema)
            });
            source.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = listContentTypeId,
                Name = "Case document",
                ParentId = siteContentTypeId,
                FieldLinks = new List<ListContentTypeFieldLinkSnapshot>
                {
                    new ListContentTypeFieldLinkSnapshot { FieldId = fieldId, InternalName = "CaseNumber" }
                }
            });
            source.SiteContentTypes.Add(new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceScope = "/sites/source",
                ContentTypeId = siteContentTypeId,
                Name = "Case document",
                Hidden = true,
                ReadOnly = true,
                Sealed = true,
                ParentContentTypeId = "0x0101",
                ParentContentTypeName = "Document",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = fieldId, Name = "CaseNumber", Role = FieldSchemaRole.DirectBinding }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot>
                {
                    new FieldSchemaSnapshot
                    {
                        Id = fieldId,
                        InternalName = "CaseNumber",
                        Title = "Case number",
                        TypeAsString = "Text",
                        SchemaXml = fieldSchema,
                        SchemaXmlSha256 = MigrationDigest.ComputeSha256(fieldSchema),
                        PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(fieldSchema),
                        Role = FieldSchemaRole.DirectBinding
                    }
                }
            });

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var listPlan = plan.Lists.Single();

            Assert.IsFalse(ContentTypeRuntimeCatalog.IsTargetRuntime(siteContentTypeId));
            Assert.IsTrue(plan.IsExecutable, string.Join(Environment.NewLine, listPlan.Issues.Select(value => value.Message)));
            Assert.AreEqual(1, listPlan.SiteContentTypes.Count);
            Assert.AreEqual(siteContentTypeId, listPlan.SiteContentTypes[0].Schema.ContentTypeId);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.Hidden);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.ReadOnly);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.Sealed);
            Assert.AreEqual("https://target.sharepoint.com/sites/target", listPlan.SiteContentTypes[0].TargetOwnerWebUrl);
        }

        [TestMethod]
        public void ListSemanticDigestExcludesMutableTargetAnalysis()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            source.SiteContentTypes.Add(new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceScope = "/sites/source",
                ContentTypeId = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                Name = "Custom item",
                ParentContentTypeId = "0x01",
                ParentContentTypeName = "Item"
            });
            source.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB",
                Name = "Custom item",
                ParentId = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            });
            var planSet = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var listPlan = planSet.Lists.Single();
            var before = listPlan.PlanDigest;
            listPlan.SiteContentTypes[0].DeferredUntilTopologyMaterialization = true;
            listPlan.SiteContentTypes[0].TargetProbe = new ContentTypeTargetProbe { ContentTypeId = "changed" };
            listPlan.SiteContentTypes[0].TargetAdmission = new ContentTypeTargetAdmission
            {
                IsEligible = true,
                Disposition = ContentTypeMaterializationDisposition.CreateOwned
            };
            listPlan.TargetProbe = new ListTargetProbe
            {
                TargetWebExists = true,
                Disposition = ListMaterializationDisposition.ReuseOwned
            };
            ListMigrationPlanFactory.SealTargetAnalysis(planSet);

            Assert.AreEqual(ListMaterializationDisposition.ReuseOwned, listPlan.Disposition);
            Assert.AreEqual(before, ListMigrationPlanFactory.ComputePlanDigest(listPlan));
            Assert.AreEqual(planSet.PlanDigest, ListMigrationPlanFactory.ComputeSetDigest(planSet));
        }

        [TestMethod]
        public void ReadOnlyRuntimeListFieldsAreRequiredButTheirValuesAreNotReplayed()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            const string schema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Modified' DisplayName='Modified' Type='DateTime' ReadOnly='TRUE' SourceID='http://schemas.microsoft.com/sharepoint/v3' />";
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                InternalName = "Modified",
                Title = "Modified",
                TypeAsString = "DateTime",
                SchemaXml = schema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schema),
                ReadOnly = true,
                FromBaseType = true
            });
            source.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 1,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "Modified", Kind = ListItemValueKind.DateTime, ScalarValue = "2026-08-31T00:00:00.0000000Z" }
                }
            });
            source.SourceItemCount = 1;

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);

            Assert.AreEqual(ListFieldMaterializationDisposition.RequireTargetRuntime, plan.Lists.Single().Fields.Single().Disposition);
        }

        [TestMethod]
        public void RuntimeListFieldCompatibilityPreservesScalarAndCollectionShapes()
        {
            Assert.IsTrue(ListFieldTypeCompatibility.IsCompatibleRuntimeType("Note", "Text"));
            Assert.IsTrue(ListFieldTypeCompatibility.IsCompatibleRuntimeType("Choice", "Text"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("MultiChoice", "Choice"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("UserMulti", "User"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("LookupMulti", "Lookup"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("TaxonomyFieldTypeMulti", "TaxonomyFieldType"));
        }

        [TestMethod]
        public void CalculatedListFieldsArePlannedInFormulaDependencyOrder()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            var alphaSchema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Alpha' DisplayName='Alpha Result' Type='Calculated' ReadOnly='TRUE'><Formula>=[Zulu Result]+1</Formula></Field>";
            var zuluSchema = "<Field ID='{55555555-5555-5555-5555-555555555555}' Name='Zulu' DisplayName='Zulu Result' Type='Calculated' ReadOnly='TRUE'><Formula>=1</Formula></Field>";
            source.Fields.Add(CalculatedListField(Guid.Parse("44444444-4444-4444-4444-444444444444"), "Alpha", "Alpha Result", alphaSchema));
            source.Fields.Add(CalculatedListField(Guid.Parse("55555555-5555-5555-5555-555555555555"), "Zulu", "Zulu Result", zuluSchema));

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var fields = plan.Lists.Single().Fields;

            Assert.AreEqual("Zulu", fields[0].InternalName);
            Assert.AreEqual("Alpha", fields[1].InternalName);
        }

        private static ListDependencySnapshot CreateListSnapshot(Guid siteId, Guid webId, Guid listId, string title)
        {
            return new ListDependencySnapshot
            {
                SourceSiteId = siteId,
                SourceWebId = webId,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceListId = listId,
                Title = title,
                BaseTemplate = 100,
                BaseType = "GenericList",
                RootFolderServerRelativeUrl = "/sites/source/Lists/" + title,
                Availability = EvidenceAvailability.Captured
            };
        }

        private static ListFieldSnapshot CalculatedListField(Guid id, string internalName, string title, string schema)
        {
            return new ListFieldSnapshot
            {
                Id = id,
                InternalName = internalName,
                Title = title,
                TypeAsString = "Calculated",
                SchemaXml = schema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schema),
                ReadOnly = true
            };
        }

        private static TopologyPlan CreateTopology(Guid siteId, Guid webId)
        {
            var plan = new TopologyPlan
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
            plan.PlanDigest = TopologyPlanner.ComputeDigest(plan);
            return plan;
        }

        private static TopologyTargetAnalysis CreateAdmittedTopologyAnalysis(TopologyPlan topology, Guid siteId, Guid webId)
        {
            return new TopologyTargetAnalysis
            {
                TopologyPlanDigest = topology.PlanDigest,
                SiteCollections = new List<TopologySiteTargetProbe>
                {
                    new TopologySiteTargetProbe
                    {
                        SourceSiteId = siteId,
                        TargetSiteCollectionUrl = "https://target.sharepoint.com/sites/target",
                        Exists = true,
                        Disposition = TopologyMaterializationDisposition.ReuseApprovedHost,
                        Webs = new List<TopologyWebTargetProbe>
                        {
                            new TopologyWebTargetProbe
                            {
                                SourceSiteId = siteId,
                                SourceWebId = webId,
                                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                                Exists = true,
                                Disposition = TopologyMaterializationDisposition.ReuseApprovedHost
                            }
                        }
                    }
                }
            };
        }

        private static PublishingPageCaptureBundle CreateSnapshot()
        {
            var fileUniqueId = Guid.Parse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
            var modifiedUtc = new DateTime(2026, 8, 26, 1, 2, 3, DateTimeKind.Utc);
            var pageBytes = Encoding.UTF8.GetBytes(
                "<%@ Page Language=\"C#\" Inherits=\"Microsoft.SharePoint.Publishing.TemplateRedirectionPage, Microsoft.SharePoint.Publishing\" %>");
            var snapshot = new PublishingPageCaptureBundle
            {
                CapturePolicy = new PageCaptureOptions
                {
                    SourcePageServerRelativeUrl = "/sites/source/Pages/source.aspx"
                },
                Source = new PageIdentity
                {
                    WebUrl = "https://source.sharepoint.com/sites/source",
                    WebServerRelativeUrl = "/sites/source",
                    PageServerRelativeUrl = "/sites/source/Pages/source.aspx",
                    FileUniqueId = fileUniqueId,
                    ContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                    ContentTypeName = "Enterprise Wiki Page",
                    Title = "Source",
                    Length = pageBytes.LongLength,
                    ModifiedUtc = modifiedUtc,
                    VersionLabel = "0.1"
                },
                PageArtifact = new PageArtifactSnapshot
                {
                    FileUniqueId = fileUniqueId,
                    ServerRelativeUrl = "/sites/source/Pages/source.aspx",
                    Bytes = MigrationArtifact.Describe(pageBytes, "application/vnd.ms-aspx", "source.aspx"),
                    ContentBase64 = Convert.ToBase64String(pageBytes),
                    PageDirective = PageDirectiveParser.Parse(Encoding.UTF8.GetString(pageBytes)),
                    Availability = EvidenceAvailability.Captured
                },
                Layout = CreateStockLayout(),
                PublishingPageContent = "<p>source</p>",
                PublishingPageContentSha256 = PublishingPageDigest.ComputeSha256("<p>source</p>"),
                Fields = new List<PageFieldValueSnapshot>
                {
                    new PageFieldValueSnapshot
                    {
                        Id = Guid.Parse("20d0d2ea-fd8e-4e91-a549-70a48e8932ef"),
                        InternalName = "OOCLReference",
                        Title = "Custom recovery field",
                        TypeAsString = "Text",
                        SchemaXml = "<Field Name=\"OOCLReference\" Type=\"Text\" />",
                        HasValue = true,
                        Kind = PageFieldValueKind.Unsupported,
                        RawType = "Contoso.CustomFieldValue",
                        RawValue = "OOCL-42",
                        RawValueJson = "{\"reference\":\"OOCL-42\"}",
                        CaptureStatus = PageCaptureStatus.CapturedWithLimitations
                    }
                },
                Security = new PageSecuritySnapshot(),
                Lifecycle = new PageLifecycleSnapshot
                {
                    CheckOutType = "Online",
                    Level = "Draft",
                    ModerationStatus = 3
                },
                SourceFence = new SourcePageFence
                {
                    FileUniqueId = fileUniqueId,
                    VersionLabel = "0.1",
                    Length = pageBytes.LongLength,
                    ModifiedUtc = modifiedUtc
                }
            };
            snapshot.Runtime = PageRuntimeResolver.Resolve(
                snapshot.PageArtifact,
                snapshot.Layout.PageDirective,
                snapshot.Source.ContentTypeId);
            snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(
                snapshot.Source,
                snapshot.Layout,
                snapshot.Fields);
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            return snapshot;
        }

        private static PageFieldValueSnapshot CreateDanglingTaxonomyField()
        {
            var fieldId = Guid.Parse("e1a5b98c-dd71-426d-acb6-e478c7a5882f");
            var storeId = Guid.Parse("11111111-2222-3333-4444-555555555555");
            var setId = Guid.Parse("66666666-7777-8888-9999-aaaaaaaaaaaa");
            var termId = Guid.Parse("bbbbbbbb-cccc-dddd-eeee-ffffffffffff");
            var catchAllData = string.Join("|",
                Convert.ToBase64String(storeId.ToByteArray()),
                Convert.ToBase64String(setId.ToByteArray()),
                Convert.ToBase64String(termId.ToByteArray()));
            var hidden = new TaxonomyHiddenListEntrySnapshot
            {
                WssId = 42,
                TermStoreId = storeId,
                TermSetId = setId,
                TermId = termId,
                Title = "Retired category",
                CatchAllData = catchAllData,
                CatchAllDataLabel = "Retired category",
                Terms = new List<TaxonomyLocalizedTextSnapshot>
                {
                    new TaxonomyLocalizedTextSnapshot
                    {
                        FieldInternalName = "Term1033",
                        Value = "Retired category"
                    }
                },
                Paths = new List<TaxonomyLocalizedTextSnapshot>
                {
                    new TaxonomyLocalizedTextSnapshot
                    {
                        FieldInternalName = "Path1033",
                        Value = "Retired category"
                    }
                }
            };
            var field = new PageFieldValueSnapshot
            {
                Id = fieldId,
                InternalName = "Wiki_x0020_Page_x0020_Categories",
                Title = "Wiki Categories",
                TypeAsString = "TaxonomyFieldTypeMulti",
                SchemaXml = "<Field Name=\"Wiki_x0020_Page_x0020_Categories\" Type=\"TaxonomyFieldTypeMulti\" />",
                HasValue = true,
                Kind = PageFieldValueKind.TaxonomyCollection,
                CaptureStatus = PageCaptureStatus.Captured,
                TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
                {
                    FieldId = fieldId,
                    FieldInternalName = "Wiki_x0020_Page_x0020_Categories",
                    TermStoreId = storeId,
                    BoundTermSetId = setId,
                    TextFieldId = Guid.Parse("12345678-1234-1234-1234-123456789012")
                },
                TaxonomyValues = new List<PageTaxonomyValueSnapshot>
                {
                    new PageTaxonomyValueSnapshot
                    {
                        Label = "Retired category",
                        TermGuid = termId.ToString("D"),
                        WssId = 42,
                        Relationship = new TaxonomyValueRelationshipSnapshot
                        {
                            CapturedAtUtc = new DateTimeOffset(2026, 8, 31, 1, 2, 3, TimeSpan.Zero),
                            State = TaxonomyRelationshipState.DanglingTermAbsent,
                            ValueHiddenListEntry = hidden,
                            TaxCatchAllHiddenListEntry = hidden
                        }
                    }
                }
            };
            PageTaxonomyRelationshipProof.Seal(field);
            return field;
        }

        private static void AddDanglingTaxonomyPlan(
            PublishingPageMigrationPackage package,
            PageFieldValueSnapshot field)
        {
            var value = field.TaxonomyValues.Single();
            var termId = Guid.Parse(value.TermGuid);
            var targetStoreId = Guid.Parse("aaaaaaaa-1111-2222-3333-bbbbbbbbbbbb");
            var targetSetId = Guid.Parse("cccccccc-1111-2222-3333-dddddddddddd");
            package.Snapshot.Fields.Add(field);
            package.Snapshot.ProfileSignals = PublishingPageProfileSignalProjector.Project(
                package.Snapshot.Source,
                package.Snapshot.Layout,
                package.Snapshot.Fields);
            package.Snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(package.Snapshot);
            package.SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(package.Snapshot);
            package.Plan.SourceSnapshotDigest = package.SnapshotDigest;
            package.Plan.PlanningPolicy.TaxonomySchemaMappings.Add(new TaxonomyTargetMapping
            {
                SourceTermStoreId = field.TaxonomyBinding.TermStoreId,
                SourceTermSetId = field.TaxonomyBinding.BoundTermSetId,
                TargetTermStoreId = targetStoreId,
                TargetTermSetId = targetSetId
            });
            package.Plan.FieldActions.Add(new PageFieldAction
            {
                SourceInternalName = field.InternalName,
                TargetInternalName = field.InternalName,
                TargetTypeAsString = field.TypeAsString,
                Disposition = PageFieldDisposition.ApplyTaxonomyRelationships,
                Reason = "Reproduce exact relationship."
            });
            package.Plan.TaxonomyRelationshipActions.Add(new TaxonomyRelationshipAction
            {
                SourceFieldId = field.Id,
                SourceFieldInternalName = field.InternalName,
                SourceTermId = termId,
                SourceWssId = value.WssId,
                SourceEvidenceSha256 = value.Relationship.EvidenceSha256,
                SourceState = TaxonomyRelationshipState.DanglingTermAbsent,
                Disposition = TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent,
                TargetFieldId = field.Id,
                TargetTextFieldId = Guid.Parse("dddddddd-1111-2222-3333-eeeeeeeeeeee"),
                TargetFieldOpen = field.TaxonomyBinding.Open,
                TargetTermStoreId = targetStoreId,
                TargetBoundTermSetId = targetSetId,
                TargetValueHiddenListTermSetId = targetSetId,
                TargetTaxCatchAllHiddenListTermSetId = targetSetId,
                Reason = "Keep the Term absent and reproduce the dangling relationship.",
                VerificationAssertions = new List<string>
                {
                    "The Term remains absent.",
                    "The target-local WssId resolves to the sealed hidden identity."
                }
            });
            package.Plan.IngredientActions = PublishingPageIngredientActionProjector.Project(package.Snapshot, package.Plan);
            var evaluation = PageIngredientPlanEvaluator.Evaluate(
                package.Snapshot.IngredientGraph,
                package.Plan.IngredientActions);
            package.Plan.MigrationOutcome = evaluation.Outcome;
            package.Plan.IngredientIssues = evaluation.Issues;
            package.State = PublishingPagePackageState.ApprovalReady;
            package.PlanDigest = PublishingPageDigest.ComputePlanDigest(package.Plan);
        }

        private static PublishingPageLayoutSnapshot CreateStockLayout()
        {
            var bytes = Encoding.UTF8.GetBytes("<%@ Page %><PublishingWebControls:RichHtmlField FieldName=\"PublishingPageContent\" runat=\"server\" />");
            return new PublishingPageLayoutSnapshot
            {
                EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                Url = "https://source.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx",
                ServerRelativeUrl = "/_catalogs/masterpage/EnterpriseWiki.aspx",
                FileName = "EnterpriseWiki.aspx",
                CustomizedPageStatus = 1,
                AssociatedContentTypeName = "Enterprise Wiki Page",
                AssociatedContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                Bytes = MigrationArtifact.Describe(bytes, "application/vnd.ms-aspx", "EnterpriseWiki.aspx"),
                ContentBase64 = Convert.ToBase64String(bytes),
                Controls = new List<PublishingPageLayoutControl>
                {
                    new PublishingPageLayoutControl
                    {
                        TagPrefix = "PublishingWebControls",
                        ControlName = "RichHtmlField",
                        FieldName = "PublishingPageContent"
                    }
                }
            };
        }

        private static PublishingPageLayoutSnapshot CreateCustomLayout()
        {
            const string authoredReference = "~site/SiteAssets/Contoso/app.js";
            var layoutBytes = Encoding.UTF8.GetBytes(
                "<%@ Page %><PublishingWebControls:TextField FieldName=\"Activity\" runat=\"server\" /><script src=\""
                + authoredReference
                + "\"></script>");
            var resourceBytes = Encoding.UTF8.GetBytes("console.log('source');");
            var fieldId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var field = Field(
                fieldId,
                "Activity",
                "Text",
                "<Field ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Activity\" DisplayName=\"Activity\" Type=\"Text\" />");
            var reference = new PublishingPageLayoutResourceReference
            {
                Attribute = "src",
                Value = authoredReference
            };
            return new PublishingPageLayoutSnapshot
            {
                EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                Url = "https://source.sharepoint.com/sites/source/_catalogs/masterpage/Custom.aspx",
                ServerRelativeUrl = "/sites/source/_catalogs/masterpage/Custom.aspx",
                FileName = "Custom.aspx",
                CustomizedPageStatus = 0,
                AssociatedContentTypeName = "Custom Publishing Page",
                AssociatedContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC001122",
                Bytes = MigrationArtifact.Describe(layoutBytes, "application/vnd.ms-aspx", "Custom.aspx"),
                ContentBase64 = Convert.ToBase64String(layoutBytes),
                Controls = new List<PublishingPageLayoutControl>
                {
                    new PublishingPageLayoutControl
                    {
                        TagPrefix = "PublishingWebControls",
                        ControlName = "TextField",
                        FieldName = "Activity"
                    }
                },
                ResourceReferences = new List<PublishingPageLayoutResourceReference> { reference },
                ResourceArtifacts = new List<PublishingPageLayoutResourceSnapshot>
                {
                    new PublishingPageLayoutResourceSnapshot
                    {
                        Reference = reference,
                        EvidenceState = PublishingPageLayoutResourceEvidenceState.Readable,
                        ResolvedSourceUrl = "https://source.sharepoint.com/sites/source/SiteAssets/Contoso/app.js",
                        Artifact = MigrationArtifact.Describe(resourceBytes, "text/javascript", "app.js"),
                        ContentBase64 = Convert.ToBase64String(resourceBytes)
                    }
                },
                AssociatedContentTypeSchema = new ContentTypeSchemaSnapshot
                {
                    EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                    Availability = EvidenceAvailability.Captured,
                    SourceWebUrl = "https://source.sharepoint.com/sites/source",
                    ContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC001122",
                    Name = "Custom Publishing Page",
                    Description = "Custom page schema",
                    Group = "Custom Content Types",
                    ParentContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC",
                    ParentContentTypeName = "Page",
                    RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                    {
                        new ContentTypeFieldLinkSnapshot
                        {
                            FieldId = fieldId,
                            Name = "Activity",
                            Role = FieldSchemaRole.DirectBinding
                        }
                    },
                    RequiredFieldClosure = new List<FieldSchemaSnapshot> { field }
                }
            };
        }

        private static PublishingPageLayoutTargetProbe CreateEligibleCustomLayoutProbe(
            PublishingPageLayoutMaterializationPlan plan)
        {
            return new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                FileExists = false,
                CanAddAndCustomizePages = true,
                Availability = EvidenceAvailability.Captured,
                ContentTypeSchema = new ContentTypeTargetProbe
                {
                    ContentTypeId = plan.ContentTypeSchema.ContentTypeId,
                    ParentContentTypeAvailable = true,
                    ResolvedParentContentTypeId = plan.ContentTypeSchema.ParentContentTypeId,
                    CanManageContentTypes = true,
                    Availability = EvidenceAvailability.Captured
                },
                Resources = plan.ResourceMaterializations
                    .Where(value => value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned)
                    .Select(value => new PublishingPageLayoutResourceTargetProbe
                    {
                        TargetServerRelativeUrl = value.TargetServerRelativeUrl,
                        FileExists = false,
                        CanWrite = true,
                        Availability = EvidenceAvailability.Captured
                    })
                    .ToList()
            };
        }

        private static FieldSchemaSnapshot Field(Guid id, string name, string type, string schemaXml)
        {
            return new FieldSchemaSnapshot
            {
                Id = id,
                InternalName = name,
                Title = name,
                TypeAsString = type,
                SchemaXml = schemaXml,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schemaXml),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schemaXml),
                Role = FieldSchemaRole.DirectBinding
            };
        }

        private static PublishingPageMigrationPackage CreateMigrationPackage()
        {
            var snapshot = CreateSnapshot();
            var snapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);
            var layoutPlan = PublishingPageLayoutPlanFactory.Create(
                snapshot.Layout,
                new Uri(snapshot.Source.WebUrl),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");
            var layoutProbe = new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = layoutPlan.TargetServerRelativeUrl,
                FileExists = true,
                ExistingAssociatedContentTypeName = "Enterprise Wiki Page",
                ExistingAssociatedContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                AssociatedContentTypeAvailable = true,
                ResolvedAssociatedContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                Availability = EvidenceAvailability.Captured
            };
            var layoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(layoutPlan, layoutProbe);
            var plan = new PublishingPageMigrationPlan
            {
                SourceSnapshotDigest = snapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                TargetWebServerRelativeUrl = "/sites/target",
                TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx",
                PageLayoutName = "EnterpriseWiki",
                TargetLifecycle = PublishingPageTargetLifecycle.Draft,
                LifecycleReason = "The source file level is 'Draft', so the target will remain Draft.",
                PlanningPolicy = new PagePlanningOptions
                {
                    TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx"
                },
                TargetProbe = new PublishingPageTargetSnapshot
                {
                    WebUrl = "https://target.sharepoint.com/sites/target",
                    WebServerRelativeUrl = "/sites/target",
                    PagesLibraryServerRelativeUrl = "/sites/target/Pages",
                    PagesLibraryBaseTemplate = 850,
                    PageContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                    PageLayoutUrl = "https://target.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx",
                    PageLayoutExists = true
                },
                LayoutMaterialization = layoutPlan,
                LayoutTargetProbe = layoutProbe,
                LayoutAdmission = layoutAdmission,
                FieldActions = new List<PageFieldAction>
                {
                    new PageFieldAction
                    {
                        SourceInternalName = "OOCLReference",
                        TargetInternalName = "OOCLReference",
                        Disposition = PageFieldDisposition.EvidenceOnly,
                        Reason = "The field is retained for a future mapper."
                    }
                },
                ExpectedPublishingPageContentSha256 = snapshot.PublishingPageContentSha256,
                RuntimeVerification = new RuntimeVerificationManifest
                {
                    Requirements = new List<RuntimeVerificationRequirement>
                    {
                        new RuntimeVerificationRequirement
                        {
                            Id = "authored-dom-equality",
                            Kind = RuntimeVerificationRequirementKind.AuthoredDomEquality,
                            Required = true,
                            Description = "Normalized authored DOM is equal."
                        }
                    }
                }
            };
            plan.IngredientActions = PublishingPageIngredientActionProjector.Project(snapshot, plan);
            var ingredientEvaluation = PageIngredientPlanEvaluator.Evaluate(snapshot.IngredientGraph, plan.IngredientActions);
            plan.MigrationOutcome = ingredientEvaluation.Outcome;
            plan.IngredientIssues = ingredientEvaluation.Issues;
            return new PublishingPageMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = DateTimeOffset.UtcNow.AddMinutes(-1),
                State = PublishingPagePackageState.ApprovalReady,
                Selection = CreateSelection(),
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(CreateSelection()),
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = snapshotDigest,
                PlanDigest = PublishingPageDigest.ComputePlanDigest(plan),
                Report = new PublishingPageMigrationReport
                {
                    Summary = "Test report"
                }
            };
        }

        private static PublishingPageWorkflowSelection CreateSelection()
        {
            return new PublishingPageWorkflowSelection
            {
                WorkflowId = EnterpriseWikiV1CohortPolicy.CohortId,
                ValidationCohort = new ValidationCohortAssessment
                {
                    CohortId = EnterpriseWikiV1CohortPolicy.CohortId,
                    PolicyVersion = EnterpriseWikiV1CohortPolicy.PolicyVersion,
                    Disposition = ValidationCohortDisposition.Included,
                    Reasons = new List<string>
                    {
                        "Enterprise Wiki Content Type lineage is included by the EW-v1 validation policy."
                    }
                }
            };
        }
    }
}
