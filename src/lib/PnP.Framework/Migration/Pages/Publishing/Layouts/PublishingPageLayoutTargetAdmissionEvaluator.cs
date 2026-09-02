using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public static class PublishingPageLayoutTargetAdmissionEvaluator
    {
        public static PublishingPageLayoutTargetAdmission Evaluate(
            PublishingPageLayoutMaterializationPlan plan,
            PublishingPageLayoutTargetProbe probe)
        {
            var issues = new List<MigrationIssue>();
            var warnings = new List<string>();
            if (plan == null)
            {
                issues.Add(Issue("SourcePageLayoutPlanUnavailable", "source-page-layout", "The source Page Layout has no materialization plan."));
                return Result(issues, warnings, PublishingPageLayoutMaterializationDisposition.Block, null);
            }

            if (plan.Disposition == PublishingPageLayoutMaterializationDisposition.Block)
            {
                issues.Add(Issue("SourcePageLayoutMaterializationBlocked", $"source-page-layout:{plan.SourceServerRelativeUrl}", plan.Reason));
                return Result(issues, warnings, PublishingPageLayoutMaterializationDisposition.Block, null);
            }

            if (probe == null || probe.Availability == EvidenceAvailability.Unavailable || probe.Availability == EvidenceAvailability.Conflict)
            {
                issues.Add(Issue("TargetPageLayoutProbeUnavailable", "target-page-layout-probe",
                    $"The target Page Layout probe is {(probe == null ? "missing" : probe.Availability.ToString())}."));
                return Result(issues, warnings, PublishingPageLayoutMaterializationDisposition.Block, null);
            }

            if (!string.Equals(plan.TargetServerRelativeUrl, probe.TargetServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                issues.Add(Issue("TargetPageLayoutProbeMismatch", "target-page-layout-probe",
                    $"The target probe path '{probe.TargetServerRelativeUrl}' does not match sealed path '{plan.TargetServerRelativeUrl}'."));
            }

            if (plan.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock)
            {
                if (!probe.FileExists)
                {
                    issues.Add(Issue("TargetStockPageLayoutUnavailable", $"target-page-layout:{plan.TargetServerRelativeUrl}",
                        "The reviewed target stock Page Layout does not exist."));
                }

                var expectedContentTypeId = probe.ResolvedAssociatedContentTypeId ?? plan.AssociatedContentTypeId;
                if (string.IsNullOrWhiteSpace(expectedContentTypeId)
                    || string.IsNullOrWhiteSpace(probe.ExistingAssociatedContentTypeId))
                {
                    issues.Add(Issue("TargetStockPageLayoutAssociationUnavailable", $"target-page-layout:{plan.TargetServerRelativeUrl}",
                        "The reviewed target stock Page Layout does not expose a resolvable associated Content Type ID."));
                }
                else if (!string.Equals(
                    probe.ExistingAssociatedContentTypeId,
                    expectedContentTypeId,
                    StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(Issue("TargetStockPageLayoutAssociationMismatch", $"target-page-layout:{plan.TargetServerRelativeUrl}",
                        $"The target stock Page Layout is associated with '{probe.ExistingAssociatedContentTypeId}', not approved Content Type '{expectedContentTypeId}'."));
                }

                return Result(issues, warnings, PublishingPageLayoutMaterializationDisposition.ReuseTargetStock, null);
            }

            var contentTypeAdmission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan.ContentTypeSchema, probe.ContentTypeSchema);
            issues.AddRange(contentTypeAdmission.Issues);
            warnings.AddRange(contentTypeAdmission.Warnings);
            var schemaCanSatisfyAssociation = contentTypeAdmission.IsEligible;
            if (!schemaCanSatisfyAssociation
                && (!probe.AssociatedContentTypeAvailable || string.IsNullOrWhiteSpace(probe.ResolvedAssociatedContentTypeId)))
            {
                issues.Add(Issue("TargetLayoutAssociatedContentTypeUnavailable",
                    $"target-page-layout-content-type:{plan.AssociatedContentTypeName}",
                    $"The target has no exact or materializable associated content type for source '{plan.AssociatedContentTypeName}' ({plan.AssociatedContentTypeId})."));
            }

            if (!schemaCanSatisfyAssociation)
            {
                foreach (var field in probe.MissingFieldBindings.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    issues.Add(Issue("TargetLayoutFieldUnavailable", $"target-page-layout-field:{field}",
                        $"The target root Web does not expose required Page Layout field binding '{field}'."));
                }
            }

            foreach (var registration in UnsupportedRegistrations(plan.RequiredRegistrations))
            {
                issues.Add(Issue("TargetLayoutServerControlUnavailable",
                    $"target-page-layout-registration:{registration.TagPrefix}",
                    $"The Page Layout requires non-platform server control namespace '{registration.Namespace}' from assembly '{registration.Assembly}'."));
            }

            foreach (var resource in plan.ResourceMaterializations.Where(value =>
                         value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.Block))
            {
                var code = resource.SourceEvidenceState == PublishingPageLayoutResourceEvidenceState.AccessDenied
                    ? "SourceLayoutResourceAccessDenied"
                    : "SourceLayoutResourceEvidenceUnavailable";
                issues.Add(Issue(code, $"page-layout-resource:{resource.SourceReference}", resource.Reason));
            }

            var resourceProbes = probe.Resources.ToDictionary(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase);
            foreach (var resource in plan.ResourceMaterializations.Where(value =>
                         value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned))
            {
                PublishingPageLayoutResourceTargetProbe resourceProbe;
                if (string.IsNullOrWhiteSpace(resource.TargetServerRelativeUrl)
                    || !resourceProbes.TryGetValue(resource.TargetServerRelativeUrl, out resourceProbe)
                    || resourceProbe.Availability == EvidenceAvailability.Unavailable
                    || resourceProbe.Availability == EvidenceAvailability.Conflict)
                {
                    issues.Add(Issue("TargetLayoutResourceProbeUnavailable",
                        $"target-page-layout-resource:{resource.TargetServerRelativeUrl}",
                        "The target layout-resource probe is missing or unavailable."));
                    continue;
                }

                if (resourceProbe.FileExists
                    && (resource.SourceArtifact == null
                        || !string.Equals(resourceProbe.ExistingBytesSha256, resource.SourceArtifact.Sha256, StringComparison.OrdinalIgnoreCase)))
                {
                    issues.Add(Issue("TargetLayoutResourceCollision",
                        $"target-page-layout-resource:{resource.TargetServerRelativeUrl}",
                        "The target asset path already exists with different bytes."));
                }
                else if (!resourceProbe.FileExists && !resourceProbe.CanWrite)
                {
                    issues.Add(Issue("TargetLayoutResourceWriteUnavailable",
                        $"target-page-layout-resource:{resource.TargetServerRelativeUrl}",
                        "The target cannot create the required corresponding-location layout asset."));
                }
            }

            if (!probe.FileExists && !probe.CanAddAndCustomizePages)
            {
                issues.Add(Issue("TargetLayoutWritePolicyUnavailable", "target-page-layout-policy:AddAndCustomizePages",
                    "The effective target permission mask denies AddAndCustomizePages required for create-only Page Layout materialization."));
            }

            var disposition = PublishingPageLayoutMaterializationDisposition.CreateOwned;
            if (probe.FileExists)
            {
                var digestMatches = plan.TargetBytes != null
                    && string.Equals(probe.ExistingBytesSha256, plan.TargetBytes.Sha256, StringComparison.OrdinalIgnoreCase);
                var desiredContentTypeId = schemaCanSatisfyAssociation && plan.ContentTypeSchema != null
                    ? plan.ContentTypeSchema.ContentTypeId
                    : probe.ResolvedAssociatedContentTypeId;
                var associationMatches = string.Equals(
                        probe.ExistingAssociatedContentTypeName,
                        plan.AssociatedContentTypeName,
                        StringComparison.OrdinalIgnoreCase)
                    || (!string.IsNullOrWhiteSpace(desiredContentTypeId)
                        && string.Equals(probe.ExistingAssociatedContentTypeId, desiredContentTypeId, StringComparison.OrdinalIgnoreCase));
                if (!digestMatches || !associationMatches)
                {
                    issues.Add(Issue("TargetPageLayoutCollision", $"target-page-layout:{plan.TargetServerRelativeUrl}",
                        $"The digest-owned target Page Layout exists with different bytes or associated content type (digestMatch={digestMatches}, associationMatch={associationMatches})."));
                }
                else
                {
                    disposition = PublishingPageLayoutMaterializationDisposition.ReuseOwned;
                    warnings.Add("The exact digest-owned target Page Layout already exists and can be reused idempotently.");
                }
            }

            return Result(issues, warnings, disposition, contentTypeAdmission);
        }

        private static IEnumerable<PublishingPageLayoutRegistration> UnsupportedRegistrations(
            IEnumerable<PublishingPageLayoutRegistration> registrations)
        {
            return (registrations ?? Array.Empty<PublishingPageLayoutRegistration>()).Where(registration =>
                registration != null
                && !string.IsNullOrWhiteSpace(registration.Assembly)
                && !registration.Assembly.StartsWith("Microsoft.SharePoint", StringComparison.OrdinalIgnoreCase)
                && !registration.Assembly.StartsWith("System.Web", StringComparison.OrdinalIgnoreCase)
                && !((registration.Namespace ?? string.Empty).StartsWith("Microsoft.", StringComparison.OrdinalIgnoreCase)
                    && registration.Assembly.IndexOf("PublicKeyToken=71e9bce111e9429c", StringComparison.OrdinalIgnoreCase) >= 0));
        }

        private static MigrationIssue Issue(string code, string ingredient, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "PublishingPageLayout",
                Ingredient = ingredient,
                Message = message
            };
        }

        private static PublishingPageLayoutTargetAdmission Result(
            IEnumerable<MigrationIssue> issues,
            IEnumerable<string> warnings,
            PublishingPageLayoutMaterializationDisposition disposition,
            ContentTypeTargetAdmission contentTypeAdmission)
        {
            var orderedIssues = issues
                .GroupBy(value => value.Code + "\u001f" + value.Ingredient + "\u001f" + value.Message, StringComparer.Ordinal)
                .Select(value => value.First())
                .OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Ingredient, StringComparer.Ordinal)
                .ThenBy(value => value.Message, StringComparer.Ordinal)
                .ToList();
            return new PublishingPageLayoutTargetAdmission
            {
                IsEligible = orderedIssues.Count == 0,
                Disposition = orderedIssues.Count == 0 ? disposition : PublishingPageLayoutMaterializationDisposition.Block,
                ContentTypeSchema = contentTypeAdmission,
                Issues = orderedIssues,
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList()
            };
        }
    }
}
