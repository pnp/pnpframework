using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.References;
using PnP.Framework.Migration.PublishingPages.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.EnterpriseWiki
{
    internal static class EnterpriseWikiTargetInspector
    {
        public static PublishingPageTargetSnapshot Inspect(
            ClientContext context,
            string targetPagePath,
            IEnumerable<PageReferenceAction> dependencies,
            PublishingPageTargetLifecycle targetLifecycle,
            ICollection<string> blockers)
        {
            var web = context.Web;
            context.Load(web, value => value.Url, value => value.ServerRelativeUrl, value => value.WebTemplate, value => value.Configuration);
            context.Load(context.Site, site => site.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var snapshot = new PublishingPageTargetSnapshot
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                WebTemplate = web.WebTemplate,
                WebConfiguration = web.Configuration
            };

            var pages = web.GetPagesLibrary();
            if (pages == null)
            {
                blockers.Add("The target web has no publishing Pages library.");
                return snapshot;
            }

            context.Load(pages,
                list => list.BaseTemplate,
                list => list.EnableVersioning,
                list => list.EnableMinorVersions,
                list => list.EnableModeration,
                list => list.ForceCheckout,
                list => list.DraftVersionVisibility);
            context.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            context.Load(pages.ContentTypes, values => values.Include(contentType => contentType.Id, contentType => contentType.Name));
            context.ExecuteQueryRetry();

            snapshot.PagesLibraryBaseTemplate = pages.BaseTemplate;
            snapshot.PagesLibraryServerRelativeUrl = pages.RootFolder.ServerRelativeUrl;
            snapshot.EnableVersioning = pages.EnableVersioning;
            snapshot.EnableMinorVersions = pages.EnableMinorVersions;
            snapshot.EnableModeration = pages.EnableModeration;
            snapshot.ForceCheckout = pages.ForceCheckout;
            snapshot.DraftVersionVisibility = pages.DraftVersionVisibility.ToString();
            if (pages.BaseTemplate != EnterpriseWikiMigrationProfile.PublishingPagesListTemplate)
            {
                blockers.Add($"The target Pages library has base template {pages.BaseTemplate}; publishing Pages template {EnterpriseWikiMigrationProfile.PublishingPagesListTemplate} is required.");
            }

            if (targetLifecycle == PublishingPageTargetLifecycle.Draft
                && (!pages.EnableVersioning || !pages.EnableMinorVersions))
            {
                blockers.Add("The source maps to Draft, but the target Pages library cannot represent a checked-in minor draft deterministically.");
            }

            var contentType = pages.ContentTypes.FirstOrDefault(value => EnterpriseWikiMigrationProfile.IsContentType(value.Id.StringValue));
            snapshot.PageContentTypeId = contentType?.Id.StringValue;
            if (contentType == null)
            {
                blockers.Add("The Enterprise Wiki Page content type is not available in the target Pages library.");
            }

            var siteRootPath = context.Site.ServerRelativeUrl == "/"
                ? string.Empty
                : context.Site.ServerRelativeUrl.TrimEnd('/');
            var layoutPath = $"{siteRootPath}/_catalogs/masterpage/{EnterpriseWikiMigrationProfile.PageLayoutFileName}";
            var layoutFile = context.Site.RootWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(layoutPath));
            context.Load(layoutFile, file => file.Exists, file => file.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            snapshot.PageLayoutExists = layoutFile.Exists;
            snapshot.PageLayoutUrl = layoutFile.Exists
                ? new Uri(new Uri(web.Url).GetLeftPart(UriPartial.Authority) + PublishingPagePath.Encode(layoutFile.ServerRelativeUrl)).AbsoluteUri
                : null;
            if (!layoutFile.Exists)
            {
                blockers.Add($"{EnterpriseWikiMigrationProfile.PageLayoutFileName} is not available in the target site collection master page gallery.");
            }

            var expectedDirectory = pages.RootFolder.ServerRelativeUrl;
            if (!string.Equals(PublishingPagePath.GetDirectoryName(targetPagePath), expectedDirectory, StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The target page must be placed in the root of '{expectedDirectory}' for the Enterprise Wiki profile.");
            }

            snapshot.TargetPageExists = PublishingPageFileProbe.Exists(context, targetPagePath);
            if (snapshot.TargetPageExists)
            {
                blockers.Add($"Create-only target page already exists: {targetPagePath}");
            }

            foreach (var path in dependencies
                         .Where(dependency => dependency.Disposition == PageReferenceDisposition.MaterializeAtTarget)
                         .Select(dependency => dependency.TargetServerRelativeUrl)
                         .Where(path => !string.IsNullOrWhiteSpace(path))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!PublishingPagePath.IsWithin(path, web.ServerRelativeUrl))
                {
                    blockers.Add($"Planned dependency target escapes the target web boundary: {path}");
                    continue;
                }

                if (PublishingPageFileProbe.Exists(context, path))
                {
                    snapshot.ExistingDependencyPaths.Add(path);
                    blockers.Add($"Create-only dependency target already exists: {path}");
                }
            }

            return snapshot;
        }
    }
}
