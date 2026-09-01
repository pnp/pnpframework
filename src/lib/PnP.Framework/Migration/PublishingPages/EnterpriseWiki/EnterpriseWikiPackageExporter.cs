using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Packaging;
using PnP.Framework.Migration.PublishingPages.References;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.EnterpriseWiki
{
    public sealed class EnterpriseWikiPackageExporter
    {
        public PublishingPageExportPackage Export(ClientContext sourceContext, PublishingPageExportOptions options)
        {
            if (sourceContext == null)
            {
                throw new ArgumentNullException(nameof(sourceContext));
            }

            ValidateOptions(options);
            var sourceWeb = sourceContext.Web;
            sourceContext.Load(sourceWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title);
            sourceContext.ExecuteQueryRetry();

            var sourcePagePath = PublishingPagePath.Normalize(sourceWeb.ServerRelativeUrl, options.SourcePageServerRelativeUrl, "Pages");
            var blockers = new List<string>();
            var warnings = new List<string>();
            var sourceCapture = PublishingPageCaptureReader.Read(sourceContext, sourcePagePath, options, blockers, warnings);
            if (!EnterpriseWikiMigrationProfile.IsContentType(sourceCapture.Identity.ContentTypeId))
            {
                blockers.Add($"Source ContentTypeId '{sourceCapture.Identity.ContentTypeId}' is not an Enterprise Wiki Page content type (Project Page is intentionally excluded).");
            }

            if (!sourceCapture.Identity.PageLayoutUrl.EndsWith("/" + EnterpriseWikiMigrationProfile.PageLayoutFileName, StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The source page uses layout '{sourceCapture.Identity.PageLayoutUrl}'. The Enterprise Wiki profile requires {EnterpriseWikiMigrationProfile.PageLayoutFileName}.");
            }

            var references = PageReferenceSnapshotReader.Read(
                sourceContext,
                sourceCapture.Identity,
                sourceCapture.PublishingPageContent,
                sourceCapture.WebParts,
                options,
                warnings);
            var afterFence = SourcePageFenceReader.Read(sourceContext, sourcePagePath);
            if (!SourcePageFenceReader.Equals(sourceCapture.SourceFence, afterFence))
            {
                blockers.Add("The source page changed while it was being exported. Discard this snapshot and export again.");
            }

            var snapshot = new PublishingPageCaptureBundle
            {
                SourceProfile = EnterpriseWikiMigrationProfile.SourceProfile,
                CapturePolicy = new PublishingPageExportOptions
                {
                    SourcePageServerRelativeUrl = sourcePagePath,
                    IncludeWebParts = options.IncludeWebParts,
                    MaximumDependencyBytes = options.MaximumDependencyBytes
                },
                Source = sourceCapture.Identity,
                PublishingPageContent = sourceCapture.PublishingPageContent,
                PublishingPageContentSha256 = PublishingPageDigest.ComputeSha256(sourceCapture.PublishingPageContent ?? string.Empty),
                Fields = sourceCapture.Fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToList(),
                WebParts = sourceCapture.WebParts
                    .OrderBy(webPart => webPart.ZoneId, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(webPart => webPart.ZoneIndex)
                    .ThenBy(webPart => webPart.Id)
                    .ToList(),
                Dependencies = references
                    .OrderBy(reference => reference.SourceAbsoluteUrl, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(reference => reference.Consumer, StringComparer.Ordinal)
                    .ToList(),
                Security = sourceCapture.Security,
                Lifecycle = sourceCapture.Lifecycle,
                SourceFence = sourceCapture.SourceFence,
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };

            return new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };
        }

        private static void ValidateOptions(PublishingPageExportOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.SourcePageServerRelativeUrl))
            {
                throw new ArgumentException("A source page path is required.", nameof(options));
            }

            if (options.MaximumDependencyBytes <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(options), "MaximumDependencyBytes must be greater than zero.");
            }
        }
    }
}
