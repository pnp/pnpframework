using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Planning
{
    internal static class PublishingPagePlanningPolicy
    {
        public static void ValidateOptions(PagePlanningOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.TargetPageServerRelativeUrl))
            {
                throw new ArgumentException("A target page path is required.", nameof(options));
            }

            if (!options.CreateOnly)
            {
                throw new NotSupportedException("Only create-page plans are supported. Deferred-field recovery remains represented by the package schema but is not executable yet.");
            }

            PagePlanningTaxonomyMappingResolver.Normalize(options);
        }

        public static void AddSnapshotDecisions(
            PublishingPageCaptureBundle snapshot,
            PagePlanningOptions options,
            ICollection<string> blockers)
        {
            if (snapshot.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments, but the selected planning policy requires inherited target permissions.");
            }

            // Managed metadata is admitted per captured relationship ingredient. A broad
            // field-level override must never turn missing or invalid evidence into a repair.
        }

        public static string DescribeLifecycleDecision(
            PageLifecycleSnapshot lifecycle,
            PublishingPageTargetLifecycle targetLifecycle,
            ICollection<string> warnings)
        {
            if (lifecycle == null || string.IsNullOrWhiteSpace(lifecycle.Level))
            {
                warnings.Add("Source lifecycle evidence is incomplete. The conservative target lifecycle is Draft.");
                return "Source lifecycle evidence is incomplete, so the target will remain Draft.";
            }

            if (string.Equals(lifecycle.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && targetLifecycle == PublishingPageTargetLifecycle.Draft)
            {
                warnings.Add("Source lifecycle evidence is contradictory. The conservative target lifecycle is Draft.");
                return $"The source reports Published but has conflicting checkout '{lifecycle.CheckOutType ?? "unknown"}' or moderation '{lifecycle.ModerationStatus?.ToString(CultureInfo.InvariantCulture) ?? "unknown"}' evidence, so the target will remain Draft.";
            }

            return targetLifecycle == PublishingPageTargetLifecycle.Published
                ? "The source file level is Published with no conflicting checkout or moderation evidence, so the target will be published."
                : $"The source file level is '{lifecycle.Level}', so the target will remain Draft.";
        }

        public static PagePlanningOptions CopyOptions(PagePlanningOptions options, string targetPagePath)
        {
            return new PagePlanningOptions
            {
                TargetPageServerRelativeUrl = targetPagePath,
                RequireInheritedPermissions = options.RequireInheritedPermissions,
                BlockOnManagedMetadata = options.BlockOnManagedMetadata,
                AllowExternalResourceReferences = options.AllowExternalResourceReferences,
                CreateOnly = options.CreateOnly,
                TaxonomySchemaMappings = (options.TaxonomySchemaMappings ?? new List<TaxonomyTargetMapping>())
                    .Select(value => new TaxonomyTargetMapping
                    {
                        SourceTermStoreId = value.SourceTermStoreId,
                        SourceTermSetId = value.SourceTermSetId,
                        TargetTermStoreId = value.TargetTermStoreId,
                        TargetTermSetId = value.TargetTermSetId,
                        Mode = value.Mode,
                        UnresolvedReferenceTargetVerifiedAbsent = value.UnresolvedReferenceTargetVerifiedAbsent,
                        UnresolvedReferenceEvidenceSha256 = value.UnresolvedReferenceEvidenceSha256
                    })
                    .ToList(),
                TaxonomyAssetMappingCatalog = PagePlanningTaxonomyMappingResolver.Clone(options.TaxonomyAssetMappingCatalog),
                TopologyPolicy = new TopologyPlanningPolicy
                {
                    DefaultChildWebTemplate = options.TopologyPolicy?.DefaultChildWebTemplate,
                    DefaultChildWebConfiguration = options.TopologyPolicy == null ? 0 : options.TopologyPolicy.DefaultChildWebConfiguration,
                    WebOverrides = (options.TopologyPolicy?.WebOverrides ?? new List<TargetWebOverride>()).Select(value => new TargetWebOverride
                    {
                        SourceWebId = value.SourceWebId,
                        TargetUrlSegment = value.TargetUrlSegment,
                        TargetTitle = value.TargetTitle,
                        TargetTemplate = value.TargetTemplate,
                        TargetConfiguration = value.TargetConfiguration
                    }).ToList()
                },
                ListTargetOverrides = (options.ListTargetOverrides ?? new List<ListTargetOverride>()).Select(value => new ListTargetOverride
                {
                    SourceListId = value.SourceListId,
                    TargetTitle = value.TargetTitle,
                    TargetRootFolderServerRelativeUrl = value.TargetRootFolderServerRelativeUrl
                }).ToList()
            };
        }
    }
}
