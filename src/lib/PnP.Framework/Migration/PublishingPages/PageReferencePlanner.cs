using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages
{
    internal static class PageReferencePlanner
    {
        public static List<PageReferenceAction> BuildActions(
            PublishingPageCaptureBundle snapshot,
            string targetWebUrl,
            string targetWebServerRelativeUrl,
            PublishingPagePlanningOptions options,
            ICollection<string> blockers)
        {
            var sourceWebUri = new Uri(UrlUtility.EnsureTrailingSlash(snapshot.Source.WebUrl));
            var targetWebUri = new Uri(UrlUtility.EnsureTrailingSlash(targetWebUrl));
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            var targetWebPath = targetWebServerRelativeUrl.TrimEnd('/');
            var result = new List<PageReferenceAction>();
            foreach (var reference in snapshot.Dependencies)
            {
                var action = new PageReferenceAction
                {
                    SnapshotDependencyId = reference.Id,
                    Disposition = PageReferenceDisposition.PreserveExternal
                };
                result.Add(action);
                if (!Uri.TryCreate(reference.SourceAbsoluteUrl, UriKind.Absolute, out var sourceUri))
                {
                    action.Disposition = PageReferenceDisposition.Block;
                    action.Diagnostics.Add("The captured dependency URL is not an absolute HTTP(S) URL.");
                    blockers.Add($"Dependency '{reference.OriginalValue}' has an invalid captured URL.");
                    continue;
                }

                if (!string.Equals(sourceWebUri.Host, sourceUri.Host, StringComparison.OrdinalIgnoreCase))
                {
                    if (reference.IsRenderableResource && !options.AllowExternalResourceReferences)
                    {
                        action.Disposition = PageReferenceDisposition.Block;
                        action.Diagnostics.Add("External renderable resources are blocked by planning policy.");
                        blockers.Add($"External resource '{sourceUri}' is blocked by policy.");
                    }

                    continue;
                }

                var sourcePath = reference.SourceServerRelativeUrl ?? Uri.UnescapeDataString(sourceUri.AbsolutePath);
                var targetPath = PublishingPagePath.IsWithin(sourcePath, sourceWebPath)
                    ? targetWebPath + sourcePath.Substring(sourceWebPath.Length)
                    : sourcePath;
                action.TargetServerRelativeUrl = targetPath;
                action.TargetAbsoluteUrl = targetWebUri.GetLeftPart(UriPartial.Authority) + PublishingPagePath.Encode(targetPath) + sourceUri.Query + sourceUri.Fragment;
                action.Disposition = PageReferenceDisposition.RewriteToTarget;
                if (!reference.IsRenderableResource || PageReferenceSnapshotReader.IsSharePointRuntimePath(sourcePath))
                {
                    continue;
                }

                if (reference.Kind == PageReferenceKind.IFrame)
                {
                    action.Disposition = PageReferenceDisposition.Block;
                    action.Diagnostics.Add("Same-tenant iframe dependencies require a separately reviewed page/application profile.");
                    blockers.Add($"Iframe dependency '{sourceUri}' is unsupported by the exact profile.");
                    continue;
                }

                if (!PublishingPagePath.IsWithin(sourcePath, sourceWebPath))
                {
                    action.Disposition = PageReferenceDisposition.Block;
                    action.Diagnostics.Add("The resource is outside the captured source web and cannot be safely materialized inside the approved target web.");
                    blockers.Add($"Same-tenant resource '{sourceUri}' is outside the source web boundary.");
                    continue;
                }

                if (reference.CaptureStatus == PageCaptureStatus.Failed
                    || string.IsNullOrWhiteSpace(reference.ContentBase64)
                    || string.IsNullOrWhiteSpace(reference.ContentSha256))
                {
                    action.Disposition = PageReferenceDisposition.Block;
                    action.Diagnostics.Add("The source payload was not captured successfully.");
                    blockers.Add($"Resource '{sourceUri}' has no restorable payload in the source snapshot.");
                    continue;
                }

                action.Disposition = PageReferenceDisposition.MaterializeAtTarget;
            }

            return result.OrderBy(action => action.SnapshotDependencyId, StringComparer.Ordinal).ToList();
        }

        public static IList<PageTextReplacement> BuildTextReplacements(
            PublishingPageIdentity source,
            string targetWebUrl,
            string targetWebServerRelativeUrl)
        {
            var sourceWebUri = new Uri(source.WebUrl);
            var targetWebUri = new Uri(targetWebUrl);
            var candidates = new[]
            {
                new PageTextReplacement
                {
                    Source = source.WebUrl.TrimEnd('/'),
                    Target = targetWebUrl.TrimEnd('/'),
                    Reason = "Map authored absolute URLs from the source web to the target web."
                },
                new PageTextReplacement
                {
                    Source = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/'),
                    Target = targetWebServerRelativeUrl.TrimEnd('/'),
                    Reason = "Map authored server-relative URLs from the source web to the target web."
                },
                new PageTextReplacement
                {
                    Source = sourceWebUri.AbsolutePath.TrimEnd('/'),
                    Target = targetWebUri.AbsolutePath.TrimEnd('/'),
                    Reason = "Map URL-encoded source web paths to the target web."
                },
                new PageTextReplacement
                {
                    Source = sourceWebUri.GetLeftPart(UriPartial.Authority),
                    Target = targetWebUri.GetLeftPart(UriPartial.Authority),
                    Reason = "Map remaining same-tenant absolute references to the target tenant origin."
                }
            };
            return candidates
                .Where(item => !string.IsNullOrEmpty(item.Source)
                    && !string.Equals(item.Source, item.Target, StringComparison.OrdinalIgnoreCase))
                .GroupBy(item => item.Source, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderByDescending(item => item.Source.Length)
                .ToList();
        }
    }
}
