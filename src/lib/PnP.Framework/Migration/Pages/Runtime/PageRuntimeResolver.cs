using PnP.Framework.Migration.Pages.Markup;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Runtime
{
    public static class PageRuntimeResolver
    {
        public static PageRuntimeSnapshot Resolve(
            PageArtifactSnapshot pageArtifact,
            PageDirectiveSnapshot layoutDirective,
            string contentTypeId)
        {
            var pageType = pageArtifact?.PageDirective?.Inherits;
            var layoutType = layoutDirective?.Inherits;
            var pageAdapter = AdapterForType(pageType);
            if (!string.Equals(pageAdapter, PageRuntimeAdapterIds.Unknown, StringComparison.Ordinal))
            {
                return Result(pageType, layoutType, pageAdapter, PageRuntimeDetectionSource.PageDirective, PageRuntimeResolutionState.Resolved);
            }

            var layoutAdapter = AdapterForType(layoutType);
            if (string.Equals(layoutAdapter, PageRuntimeAdapterIds.Publishing, StringComparison.Ordinal))
            {
                return Result(pageType, layoutType, layoutAdapter, PageRuntimeDetectionSource.LayoutDirective, PageRuntimeResolutionState.Resolved);
            }

            if (!string.IsNullOrWhiteSpace(contentTypeId)
                && (contentTypeId.StartsWith(BuiltInContentTypeId.EnterpriseWikiPage, StringComparison.OrdinalIgnoreCase)
                    || contentTypeId.StartsWith(BuiltInContentTypeId.ProjectPage, StringComparison.OrdinalIgnoreCase)))
            {
                return Result(
                    pageType,
                    layoutType,
                    PageRuntimeAdapterIds.Publishing,
                    PageRuntimeDetectionSource.ContentTypeFallback,
                    PageRuntimeResolutionState.Fallback,
                    "The CLR runtime was not declared by readable page/layout markup; the Publishing adapter was selected from Content Type evidence as a fallback.");
            }

            return Result(
                pageType,
                layoutType,
                PageRuntimeAdapterIds.Unknown,
                PageRuntimeDetectionSource.Unknown,
                PageRuntimeResolutionState.Unknown,
                "No recognized CLR page runtime was found. Content Type and layout names cannot establish an executable adapter.");
        }

        private static string AdapterForType(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return PageRuntimeAdapterIds.Unknown;
            }

            if (typeName.IndexOf("WikiEditPage", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return PageRuntimeAdapterIds.Wiki;
            }

            if (typeName.IndexOf("TemplateRedirectionPage", StringComparison.OrdinalIgnoreCase) >= 0
                || typeName.IndexOf("PublishingLayoutPage", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return PageRuntimeAdapterIds.Publishing;
            }

            if (typeName.IndexOf("WebPartPage", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return PageRuntimeAdapterIds.WebPartPage;
            }

            return PageRuntimeAdapterIds.Unknown;
        }

        private static PageRuntimeSnapshot Result(
            string pageType,
            string layoutType,
            string adapterId,
            PageRuntimeDetectionSource detectionSource,
            PageRuntimeResolutionState resolutionState,
            string diagnostic = null)
        {
            return new PageRuntimeSnapshot
            {
                PageDeclaredType = pageType,
                LayoutDeclaredType = layoutType,
                AdapterId = adapterId,
                DetectionSource = detectionSource,
                ResolutionState = resolutionState,
                Diagnostics = string.IsNullOrWhiteSpace(diagnostic)
                    ? new List<string>()
                    : new List<string> { diagnostic }
            };
        }
    }
}
