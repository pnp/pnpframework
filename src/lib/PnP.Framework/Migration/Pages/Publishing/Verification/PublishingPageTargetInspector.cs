using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.References;
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
            ICollection<string> blockers)
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
                includeListInventory: false);
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
            bool includeListInventory = false)
        {
            var web = context.Web;
            if (!web.IsPropertyAvailable("Url")
                || !web.IsPropertyAvailable("ServerRelativeUrl")
                || !web.IsPropertyAvailable("WebTemplate")
                || !web.IsPropertyAvailable("Configuration")
                || !context.Site.IsPropertyAvailable("ServerRelativeUrl"))
            {
                context.Load(web, value => value.Url, value => value.ServerRelativeUrl, value => value.WebTemplate, value => value.Configuration);
                context.Load(context.Site, site => site.ServerRelativeUrl);
                context.ExecuteQueryRetry();
            }
            var snapshot = new PublishingPageTargetSnapshot
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                WebTemplate = web.WebTemplate,
                WebConfiguration = web.Configuration
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
            if (!string.Equals(expectedDirectory, actualDirectory, StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The current Publishing Page writer requires the target page in the root of '{actualDirectory}'.");
            }

            var dependencyPaths = dependencies
                .Where(dependency => dependency.Disposition == PageReferenceDisposition.MaterializeAtTarget)
                .Select(dependency => dependency.TargetServerRelativeUrl)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            foreach (var path in dependencyPaths.Where(path => !PagePath.IsWithin(path, web.ServerRelativeUrl)))
            {
                blockers.Add($"Planned dependency target escapes the target web boundary: {path}");
            }

            var pathsToProbe = new[] { targetPagePath }
                .Concat(dependencyPaths.Where(path => PagePath.IsWithin(path, web.ServerRelativeUrl)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var files = pathsToProbe.ToDictionary(
                path => path,
                path => web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path)),
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

            snapshot.TargetPageExists = files == null
                ? PageFileProbe.Exists(context, targetPagePath)
                : files[targetPagePath].Exists;
            if (snapshot.TargetPageExists)
            {
                blockers.Add($"Create-only target page already exists: {targetPagePath}");
            }

            foreach (var path in dependencyPaths.Where(path => PagePath.IsWithin(path, web.ServerRelativeUrl)))
            {
                var exists = files == null ? PageFileProbe.Exists(context, path) : files[path].Exists;
                if (exists)
                {
                    snapshot.ExistingDependencyPaths.Add(path);
                    blockers.Add($"Create-only dependency target already exists: {path}");
                }
            }

            return snapshot;
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
