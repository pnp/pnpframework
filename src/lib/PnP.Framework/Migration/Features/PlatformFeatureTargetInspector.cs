using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Features
{
    internal static class PlatformFeatureTargetInspector
    {
        public static IDictionary<Guid, PlatformFeatureTargetProbe> Inspect(
            ClientContext anchorContext,
            IEnumerable<PlatformFeatureMaterializationPlan> plans)
        {
            if (anchorContext == null)
            {
                throw new ArgumentNullException(nameof(anchorContext));
            }

            var featurePlans = (plans ?? Enumerable.Empty<PlatformFeatureMaterializationPlan>())
                .Where(value => value != null)
                .ToArray();
            if (featurePlans.Length == 0)
            {
                return new Dictionary<Guid, PlatformFeatureTargetProbe>();
            }
            var targetWebUrl = featurePlans.Select(value => value.TargetWebUrl)
                .Distinct(StringComparer.OrdinalIgnoreCase).Single();
            try
            {
                using (var context = anchorContext.Clone(targetWebUrl))
                {
                    context.Load(context.Site.Features, values => values.Include(value => value.DefinitionId));
                    context.Load(context.Web.Features, values => values.Include(value => value.DefinitionId));
                    context.Load(context.Site.RootWeb,
                        value => value.Url,
                        value => value.EffectiveBasePermissions);
                    context.Load(context.Web,
                        value => value.Url,
                        value => value.EffectiveBasePermissions);
                    context.Load(context.Web.AvailableContentTypes, values => values.Include(value => value.Id));
                    context.ExecuteQueryRetry();

                    var activeSiteFeatureIds = new HashSet<Guid>(context.Site.Features.AsEnumerable().Select(value => value.DefinitionId));
                    var activeWebFeatureIds = new HashSet<Guid>(context.Web.Features.AsEnumerable().Select(value => value.DefinitionId));
                    var availableContentTypeIds = new HashSet<string>(
                        context.Web.AvailableContentTypes.AsEnumerable().Select(value => value.Id.StringValue),
                        StringComparer.OrdinalIgnoreCase);
                    var canActivateSiteFeature = context.Site.RootWeb.EffectiveBasePermissions.Has(PermissionKind.ManageWeb);
                    var canActivateWebFeature = context.Web.EffectiveBasePermissions.Has(PermissionKind.ManageWeb);
                    return featurePlans.ToDictionary(
                        value => value.FeatureId,
                        value => CreateProbe(
                            value,
                            activeSiteFeatureIds,
                            activeWebFeatureIds,
                            availableContentTypeIds,
                            canActivateSiteFeature,
                            canActivateWebFeature));
                }
            }
            catch (ServerException exception)
            {
                return featurePlans.ToDictionary(
                    value => value.FeatureId,
                    value => FailedProbe(value, exception.Message));
            }
        }

        public static PlatformFeatureTargetProbe DeferUntilTopologyMaterialization(PlatformFeatureMaterializationPlan plan)
        {
            return new PlatformFeatureTargetProbe
            {
                FeatureId = plan.FeatureId,
                Scope = plan.Scope,
                TargetWebUrl = plan.TargetWebUrl,
                DeferredUntilTopologyMaterialization = true,
                CanActivate = true
            };
        }

        private static PlatformFeatureTargetProbe CreateProbe(
            PlatformFeatureMaterializationPlan plan,
            ISet<Guid> activeSiteFeatureIds,
            ISet<Guid> activeWebFeatureIds,
            ISet<string> availableContentTypeIds,
            bool canActivateSiteFeature,
            bool canActivateWebFeature)
        {
            var activeFeatureIds = plan.Scope == PlatformFeatureScope.SiteCollection
                ? activeSiteFeatureIds
                : activeWebFeatureIds;
            var canActivate = plan.Scope == PlatformFeatureScope.SiteCollection
                ? canActivateSiteFeature
                : canActivateWebFeature;
            var probe = new PlatformFeatureTargetProbe
            {
                FeatureId = plan.FeatureId,
                Scope = plan.Scope,
                TargetWebUrl = plan.TargetWebUrl,
                TargetScopeExists = true,
                IsActive = activeFeatureIds.Contains(plan.FeatureId),
                CanActivate = canActivate,
                AvailableContentTypeIds = plan.ExpectedContentTypeIds
                    .Where(availableContentTypeIds.Contains)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList()
            };
            if (!probe.IsActive && !probe.CanActivate)
            {
                probe.Issues.Add(Issue(
                    "TargetFeatureActivationUnavailable",
                    plan,
                    "The target scope does not expose the required feature and the current identity does not have ManageWeb permission to activate it."));
            }
            if (probe.IsActive)
            {
                foreach (var missing in plan.ExpectedContentTypeIds.Where(value => !availableContentTypeIds.Contains(value)))
                {
                    probe.Issues.Add(Issue(
                        "ActiveTargetFeatureContractMissing",
                        plan,
                        "The target feature is active, but its required runtime content type is unavailable: " + missing + "."));
                }
            }
            return probe;
        }

        private static PlatformFeatureTargetProbe FailedProbe(PlatformFeatureMaterializationPlan plan, string message)
        {
            var probe = new PlatformFeatureTargetProbe
            {
                FeatureId = plan.FeatureId,
                Scope = plan.Scope,
                TargetWebUrl = plan.TargetWebUrl
            };
            probe.Issues.Add(Issue(
                "TargetFeatureScopeUnavailable",
                plan,
                "The target feature scope could not be inspected: " + message));
            return probe;
        }

        private static MigrationIssue Issue(string code, PlatformFeatureMaterializationPlan plan, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "platform-feature:" + plan.FeatureId.ToString("D"),
                Ingredient = "PlatformFeature.TargetProbe",
                Message = message,
                TargetIdentity = plan.TargetWebUrl + "#site-feature:" + plan.FeatureId.ToString("D")
            };
        }
    }
}
