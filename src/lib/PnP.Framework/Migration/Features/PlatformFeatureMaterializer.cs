using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Features
{
    internal static class PlatformFeatureMaterializer
    {
        public static bool Ensure(ClientContext anchorContext, PlatformFeatureMaterializationPlan plan)
        {
            if (anchorContext == null)
            {
                throw new ArgumentNullException(nameof(anchorContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (plan.Disposition == PlatformFeatureMaterializationDisposition.Block)
            {
                throw new InvalidOperationException("A blocked platform feature cannot be materialized.");
            }

            var before = PlatformFeatureTargetInspector.Inspect(anchorContext, new[] { plan })[plan.FeatureId];
            if (!before.IsAdmitted)
            {
                throw new InvalidOperationException("Fresh platform feature preflight failed: "
                    + string.Join("; ", before.Issues.Select(value => value.Message)));
            }
            if (before.IsActive)
            {
                return false;
            }

            using (var context = anchorContext.Clone(plan.TargetWebUrl))
            {
                switch (plan.Scope)
                {
                    case PlatformFeatureScope.SiteCollection:
                        context.Site.Features.Add(plan.FeatureId, true, FeatureDefinitionScope.Farm);
                        break;
                    case PlatformFeatureScope.Web:
                        context.Web.Features.Add(plan.FeatureId, true, FeatureDefinitionScope.Farm);
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported platform feature scope: " + plan.Scope + ".");
                }
                context.ExecuteQueryRetry();
            }

            var after = PlatformFeatureTargetInspector.Inspect(anchorContext, new[] { plan })[plan.FeatureId];
            if (!after.IsAdmitted || !after.IsActive)
            {
                throw new InvalidOperationException("Fresh platform feature readback failed: "
                    + string.Join("; ", after.Issues.Select(value => value.Message)));
            }
            return true;
        }
    }
}
