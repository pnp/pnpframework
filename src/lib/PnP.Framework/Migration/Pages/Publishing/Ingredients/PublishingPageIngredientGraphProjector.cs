using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal enum PublishingPageIngredientGraphProjectionRevision
    {
        LegacyV1 = 1,
        Version2 = 2,
        Version3 = 3,
        Version4 = 4,
        Version5 = 5,
        Version6 = 6,
        CurrentV7 = 7
    }

    internal static class PublishingPageIngredientGraphProjector
    {
        public const string ProjectionVersionV2 = "pnp-publishing-page-ingredient-projection/v2";

        public const string ProjectionVersionV3 = "pnp-publishing-page-ingredient-projection/v3";

        public const string ProjectionVersionV4 = "pnp-publishing-page-ingredient-projection/v4";

        public const string ProjectionVersionV5 = "pnp-publishing-page-ingredient-projection/v5";

        public const string ProjectionVersionV6 = "pnp-publishing-page-ingredient-projection/v6";

        public const string CurrentProjectionVersion = "pnp-publishing-page-ingredient-projection/v7";

        public static CanonicalPageIngredientGraph Project(PublishingPageCaptureBundle snapshot)
        {
            return Project(
                snapshot,
                PublishingPageIngredientGraphProjectionRevision.CurrentV7,
                CurrentProjectionVersion);
        }

        internal static CanonicalPageIngredientGraph ProjectLegacy(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.LegacyV1, null);
        }

        internal static CanonicalPageIngredientGraph ProjectVersion2(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version2, ProjectionVersionV2);
        }

        internal static CanonicalPageIngredientGraph ProjectVersion3(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version3, ProjectionVersionV3);
        }

        internal static CanonicalPageIngredientGraph ProjectVersion4(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version4, ProjectionVersionV4);
        }

        internal static CanonicalPageIngredientGraph ProjectVersion5(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version5, ProjectionVersionV5);
        }

        internal static CanonicalPageIngredientGraph ProjectVersion6(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version6, ProjectionVersionV6);
        }

        internal static CanonicalPageIngredientGraph ProjectCurrentUnversioned(PublishingPageCaptureBundle snapshot)
        {
            return Project(snapshot, PublishingPageIngredientGraphProjectionRevision.Version2, null);
        }

        internal static CanonicalPageIngredientGraph ProjectForVersion(
            PublishingPageCaptureBundle snapshot,
            string projectionVersion)
        {
            if (string.Equals(projectionVersion, CurrentProjectionVersion, StringComparison.Ordinal))
            {
                return Project(snapshot);
            }
            if (string.Equals(projectionVersion, ProjectionVersionV6, StringComparison.Ordinal))
            {
                return ProjectVersion6(snapshot);
            }
            if (string.Equals(projectionVersion, ProjectionVersionV5, StringComparison.Ordinal))
            {
                return ProjectVersion5(snapshot);
            }
            if (string.Equals(projectionVersion, ProjectionVersionV4, StringComparison.Ordinal))
            {
                return ProjectVersion4(snapshot);
            }
            if (string.Equals(projectionVersion, ProjectionVersionV3, StringComparison.Ordinal))
            {
                return ProjectVersion3(snapshot);
            }
            if (string.Equals(projectionVersion, ProjectionVersionV2, StringComparison.Ordinal))
            {
                return ProjectVersion2(snapshot);
            }
            throw new ArgumentException($"Unsupported Publishing Page ingredient projection '{projectionVersion}'.", nameof(projectionVersion));
        }

        internal static bool UsesTransactionDependencies(PublishingPageIngredientGraphProjectionRevision revision)
        {
            return revision == PublishingPageIngredientGraphProjectionRevision.Version4
                || revision == PublishingPageIngredientGraphProjectionRevision.Version5
                || revision == PublishingPageIngredientGraphProjectionRevision.Version6
                || revision == PublishingPageIngredientGraphProjectionRevision.CurrentV7;
        }

        internal static bool UsesOwnerWebDependencies(PublishingPageIngredientGraphProjectionRevision revision)
        {
            return revision == PublishingPageIngredientGraphProjectionRevision.Version5
                || revision == PublishingPageIngredientGraphProjectionRevision.Version6
                || revision == PublishingPageIngredientGraphProjectionRevision.CurrentV7;
        }

        private static CanonicalPageIngredientGraph Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageIngredientGraphProjectionRevision revision,
            string projectionVersion)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            var graph = new CanonicalPageIngredientGraph
            {
                ProjectionVersion = projectionVersion
            };
            PublishingPageCoreIngredientGraphProjector.Project(snapshot, graph, revision);
            PublishingPageLayoutIngredientGraphProjector.Project(snapshot, graph, revision);
            PublishingPageTopologyIngredientGraphProjector.Project(snapshot, graph, revision);
            PublishingPageWebPartIngredientGraphProjector.Project(snapshot, graph, revision);
            PublishingPageListIngredientGraphProjector.Project(snapshot, graph, revision);
            PublishingPageReferenceIngredientGraphProjector.Project(snapshot, graph, revision);
            // Captured inventories can contain the same semantic binding more than once
            // (for example, duplicate ListItem field-value evidence). Ingredient edges are
            // set-valued, so collapse only edges whose complete semantic identity matches.
            graph.Edges = graph.Edges
                .GroupBy(
                    edge => edge.FromIngredientId + "\u001f" + edge.ToIngredientId + "\u001f"
                        + edge.Relationship + "\u001f" + edge.Requirement + "\u001f" + edge.Condition,
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
            return graph;
        }
    }
}
