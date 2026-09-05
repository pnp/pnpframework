using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageTargetInspector
    {
        private const int PublishingPagesListTemplate = 850;

        public static PublishingPageTargetSnapshot Inspect(
            ClientContext context,
            string targetPagePath,
            IEnumerable<PageReferenceAction> dependencies,
            PublishingPageTargetLifecycle targetLifecycle,
            PublishingPageLayoutMaterializationPlan layoutPlan,
            PublishingPageLayoutTargetProbe layoutProbe,
            ICollection<string> blockers,
            TopologyTargetAnalysis dependencyTopology = null)
        {
            Microsoft.SharePoint.Client.List ignored;
            return Inspect(
                context,
                targetPagePath,
                dependencies,
                targetLifecycle,
                layoutPlan,
                layoutProbe,
                blockers,
                out ignored,
                includeListInventory: false,
                resolvePageCollision: false,
                pageOriginalIdentifier: null,
                dependencyTopology: dependencyTopology);
        }

        public static PublishingPageTargetSnapshot InspectForPlanning(
            ClientContext context,
            string targetPagePath,
            string pageOriginalIdentifier,
            IEnumerable<PageReferenceAction> dependencies,
            PublishingPageTargetLifecycle targetLifecycle,
            PublishingPageLayoutMaterializationPlan layoutPlan,
            PublishingPageLayoutTargetProbe layoutProbe,
            ICollection<string> blockers,
            out Microsoft.SharePoint.Client.List pages,
            bool includeListInventory = false,
            TopologyTargetAnalysis dependencyTopology = null)
        {
            return Inspect(
                context,
                targetPagePath,
                dependencies,
                targetLifecycle,
                layoutPlan,
                layoutProbe,
                blockers,
                out pages,
                includeListInventory,
                resolvePageCollision: true,
                pageOriginalIdentifier: pageOriginalIdentifier,
                dependencyTopology: dependencyTopology);
        }

        public static PublishingPageTargetSnapshot Inspect(
            ClientContext context,
            string targetPagePath,
            IEnumerable<PageReferenceAction> dependencies,
            PublishingPageTargetLifecycle targetLifecycle,
            PublishingPageLayoutMaterializationPlan layoutPlan,
            PublishingPageLayoutTargetProbe layoutProbe,
            ICollection<string> blockers,
            out Microsoft.SharePoint.Client.List pages,
            bool includeListInventory = false,
            TopologyTargetAnalysis dependencyTopology = null)
        {
            return Inspect(
                context,
                targetPagePath,
                dependencies,
                targetLifecycle,
                layoutPlan,
                layoutProbe,
                blockers,
                out pages,
                includeListInventory,
                resolvePageCollision: false,
                pageOriginalIdentifier: null,
                dependencyTopology: dependencyTopology);
        }

        private static PublishingPageTargetSnapshot Inspect(
            ClientContext context,
            string targetPagePath,
            IEnumerable<PageReferenceAction> dependencies,
            PublishingPageTargetLifecycle targetLifecycle,
            PublishingPageLayoutMaterializationPlan layoutPlan,
            PublishingPageLayoutTargetProbe layoutProbe,
            ICollection<string> blockers,
            out Microsoft.SharePoint.Client.List pages,
            bool includeListInventory,
            bool resolvePageCollision,
            string pageOriginalIdentifier,
            TopologyTargetAnalysis dependencyTopology)
        {
            var web = context.Web;
            if (!web.IsPropertyAvailable("Url")
                || !web.IsPropertyAvailable("ServerRelativeUrl")
                || !web.IsPropertyAvailable("WebTemplate")
                || !web.IsPropertyAvailable("Configuration")
                || !context.Site.IsPropertyAvailable("ServerRelativeUrl"))
            {
                context.Load(web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl, value => value.WebTemplate, value => value.Configuration);
                context.Load(context.Site, site => site.ServerRelativeUrl);
                context.ExecuteQueryRetry();
            }
            var snapshot = new PublishingPageTargetSnapshot
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                WebTemplate = web.WebTemplate,
                WebConfiguration = web.Configuration,
                PreferredTargetPageServerRelativeUrl = targetPagePath,
                TargetPageServerRelativeUrl = targetPagePath
            };

            var expectedDirectory = PagePath.GetDirectoryName(targetPagePath);
            pages = web.GetList(expectedDirectory);
            context.Load(pages,
                list => list.BaseTemplate,
                list => list.EnableVersioning,
                list => list.EnableMinorVersions,
                list => list.EnableModeration,
                list => list.ForceCheckout,
                list => list.DraftVersionVisibility);
            context.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            context.Load(pages.ContentTypes, values => values.Include(contentType => contentType.Id, contentType => contentType.Name));
            context.Load(pages.Fields, values => values.Include(
                field => field.InternalName,
                field => field.TypeAsString,
                field => field.ReadOnlyField));
            if (includeListInventory)
            {
                context.Load(web, value => value.EffectiveBasePermissions);
                context.Load(web.Lists, values => values.Include(
                    value => value.Id,
                    value => value.Title,
                    value => value.BaseTemplate,
                    value => value.RootFolder.ServerRelativeUrl,
                    value => value.RootFolder.Properties));
            }
            try
            {
                context.ExecuteQueryRetry();
            }
            catch (ServerException exception) when (IsMissing(exception))
            {
                pages = null;
                blockers.Add("The target web has no publishing Pages library.");
                return snapshot;
            }

            snapshot.PagesLibraryBaseTemplate = pages.BaseTemplate;
            snapshot.PagesLibraryServerRelativeUrl = pages.RootFolder.ServerRelativeUrl;
            snapshot.EnableVersioning = pages.EnableVersioning;
            snapshot.EnableMinorVersions = pages.EnableMinorVersions;
            snapshot.EnableModeration = pages.EnableModeration;
            snapshot.ForceCheckout = pages.ForceCheckout;
            snapshot.DraftVersionVisibility = pages.DraftVersionVisibility.ToString();
            if (pages.BaseTemplate != PublishingPagesListTemplate)
            {
                blockers.Add($"The target Pages library has base template {pages.BaseTemplate}; publishing Pages template {PublishingPagesListTemplate} is required.");
            }

            if (targetLifecycle == PublishingPageTargetLifecycle.Draft
                && (!pages.EnableVersioning || !pages.EnableMinorVersions))
            {
                blockers.Add("The source maps to Draft, but the target Pages library cannot represent a checked-in minor draft deterministically.");
            }

            snapshot.PageContentTypeId = ResolvePageContentTypeId(pages.ContentTypes, layoutPlan, layoutProbe, blockers);
            snapshot.PageLayoutExists = layoutProbe?.FileExists == true;
            snapshot.PageLayoutUrl = layoutProbe == null || string.IsNullOrWhiteSpace(layoutProbe.TargetServerRelativeUrl)
                ? null
                : new Uri(new Uri(web.Url).GetLeftPart(UriPartial.Authority) + PagePath.Encode(layoutProbe.TargetServerRelativeUrl)).AbsoluteUri;

            var actualDirectory = pages.RootFolder.ServerRelativeUrl;
            if (!PagePath.IsWithin(expectedDirectory, actualDirectory))
            {
                blockers.Add($"The target Page path '{targetPagePath}' is outside the publishing Pages library '{actualDirectory}'.");
            }

            var dependencyPaths = dependencies
                .Where(dependency => dependency.Disposition == PageReferenceDisposition.MaterializeAtTarget)
                .Select(dependency => dependency.TargetServerRelativeUrl)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var dependencyOwners = new Dictionary<string, Web>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in dependencyPaths)
            {
                var owner = ResolveDependencyOwner(context, web, path, dependencyTopology);
                if (owner == null)
                {
                    blockers.Add($"Planned dependency target is outside the reviewed target topology: {path}");
                    continue;
                }
                dependencyOwners[path] = owner;
            }

            var pathsToProbe = new[] { targetPagePath }
                .Concat(dependencyOwners.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var files = pathsToProbe.ToDictionary(
                path => path,
                path => (string.Equals(path, targetPagePath, StringComparison.OrdinalIgnoreCase)
                    ? web
                    : dependencyOwners[path]).GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var file in files.Values)
            {
                context.Load(file, value => value.Exists);
            }
            try
            {
                context.ExecuteQueryRetry();
            }
            catch (ServerException exception) when (IsMissing(exception))
            {
                // Older target farms can throw for a missing file instead of returning Exists=false.
                // Preserve the same semantics with individual probes only on that compatibility path.
                files = null;
            }

            snapshot.PreferredTargetPageExists = files == null
                ? PageFileProbe.Exists(context, targetPagePath)
                : files[targetPagePath].Exists;
            snapshot.TargetPageExists = snapshot.PreferredTargetPageExists;
            if (snapshot.PreferredTargetPageExists && resolvePageCollision)
            {
                if (string.IsNullOrWhiteSpace(pageOriginalIdentifier))
                {
                    throw new ArgumentException("A stable Page source identity is required for planning-time collision resolution.", nameof(pageOriginalIdentifier));
                }
                var targetFolder = web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(expectedDirectory));
                context.Load(targetFolder.Files, values => values.Include(value => value.ServerRelativeUrl));
                context.ExecuteQueryRetry();
                var resolution = PublishingPageTargetPathResolver.Resolve(
                    targetPagePath,
                    pageOriginalIdentifier,
                    targetFolder.Files.AsEnumerable().Select(value => value.ServerRelativeUrl));
                snapshot.TargetPageServerRelativeUrl = resolution.TargetPageServerRelativeUrl;
                snapshot.TargetPathCollisionResolved = resolution.CollisionResolved;
                snapshot.TargetPathResolutionReason = resolution.Reason;
                snapshot.TargetPageExists = targetFolder.Files.AsEnumerable().Any(value => string.Equals(
                    Uri.UnescapeDataString(value.ServerRelativeUrl),
                    Uri.UnescapeDataString(resolution.TargetPageServerRelativeUrl),
                    StringComparison.OrdinalIgnoreCase));
            }
            if (snapshot.TargetPageExists)
            {
                blockers.Add($"Create-only target page already exists: {snapshot.TargetPageServerRelativeUrl}");
            }

            foreach (var path in dependencyOwners.Keys)
            {
                var exists = files == null
                    ? FileExists(context, dependencyOwners[path], path)
                    : files[path].Exists;
                if (exists)
                {
                    snapshot.ExistingDependencyPaths.Add(path);
                    blockers.Add($"Create-only dependency target already exists: {path}");
                }
            }

            return snapshot;
        }

        private static Web ResolveDependencyOwner(
            ClientContext context,
            Web pageWeb,
            string targetPath,
            TopologyTargetAnalysis topology)
        {
            if (PagePath.IsWithin(targetPath, pageWeb.ServerRelativeUrl))
            {
                return pageWeb;
            }

            var candidate = topology?.SiteCollections
                .SelectMany(value => value.Webs ?? Array.Empty<TopologyWebTargetProbe>())
                .Where(value => value != null
                    && value.Exists
                    && value.TargetWebId.HasValue
                    && !string.IsNullOrWhiteSpace(value.TargetServerRelativeUrl)
                    && PagePath.IsWithin(targetPath, value.TargetServerRelativeUrl))
                .OrderByDescending(value => value.TargetServerRelativeUrl.Length)
                .FirstOrDefault();
            return candidate?.TargetWebId == null
                ? null
                : context.Site.OpenWebById(candidate.TargetWebId.Value);
        }

        private static bool FileExists(ClientContext context, Web owner, string path)
        {
            var file = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path));
            context.Load(file, value => value.Exists);
            try
            {
                context.ExecuteQueryRetry();
                return file.Exists;
            }
            catch (ServerException exception) when (IsMissing(exception))
            {
                return false;
            }
        }

        private static bool IsMissing(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }

        private static string ResolvePageContentTypeId(
            IEnumerable<ContentType> contentTypes,
            PublishingPageLayoutMaterializationPlan layoutPlan,
            PublishingPageLayoutTargetProbe layoutProbe,
            ICollection<string> blockers)
        {
            var expectedRootId = layoutPlan?.ContentTypeSchema?.ContentTypeId
                ?? layoutProbe?.ResolvedAssociatedContentTypeId
                ?? layoutPlan?.AssociatedContentTypeId;
            if (string.IsNullOrWhiteSpace(expectedRootId))
            {
                blockers.Add("The approved Page Layout does not resolve an associated Content Type ID.");
                return null;
            }

            var candidates = contentTypes
                .Where(value => string.Equals(value.Id.StringValue, expectedRootId, StringComparison.OrdinalIgnoreCase)
                    || (string.Equals(value.Id.GetParentIdValue(), expectedRootId, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(value.Name, layoutPlan?.AssociatedContentTypeName, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(value => value.Id.StringValue.Length)
                .ThenBy(value => value.Id.StringValue, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var exact = candidates.FirstOrDefault(value =>
                string.Equals(value.Id.StringValue, expectedRootId, StringComparison.OrdinalIgnoreCase));
            if (exact != null)
            {
                return exact.Id.StringValue;
            }

            if (candidates.Length == 1)
            {
                return candidates[0].Id.StringValue;
            }

            if (candidates.Length > 1)
            {
                blockers.Add($"The target Pages library exposes multiple Content Types derived from the Page Layout association '{expectedRootId}'; an exact target cannot be selected deterministically: {string.Join(", ", candidates.Select(value => value.Id.StringValue))}.");
                return null;
            }

            if (layoutPlan?.ContentTypeSchema != null)
            {
                return layoutPlan.ContentTypeSchema.ContentTypeId;
            }

            blockers.Add($"The Content Type associated with the approved Page Layout is unavailable in the target Pages library: {expectedRootId}.");
            return null;
        }
    }
}
