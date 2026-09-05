using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Capture
{
    internal sealed class PublishingPagePackageExporter
    {
        public PublishingPageExportPackage Export(
            ClientContext sourceContext,
            PageCaptureOptions options,
            PublishingPageWorkflowPolicy workflowPolicy)
        {
            return Export(sourceContext, options, workflowPolicy, null);
        }

        public PublishingPageExportPackage Export(
            ClientContext sourceContext,
            PageCaptureOptions options,
            PublishingPageWorkflowPolicy workflowPolicy,
            IMigrationArtifactStore artifactStore)
        {
            if (sourceContext == null)
            {
                throw new ArgumentNullException(nameof(sourceContext));
            }
            if (workflowPolicy == null)
            {
                throw new ArgumentNullException(nameof(workflowPolicy));
            }

            ValidateOptions(options);
            var sourceWeb = sourceContext.Web;
            sourceContext.Load(sourceWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title);
            sourceContext.ExecuteQueryRetry();

            var sourcePagePath = PagePath.Normalize(sourceWeb.ServerRelativeUrl, options.SourcePageServerRelativeUrl, "Pages");
            var blockers = new List<string>();
            var warnings = new List<string>();
            var sourceCapture = PublishingPageCaptureReader.Read(
                sourceContext,
                sourcePagePath,
                options,
                artifactStore,
                blockers,
                warnings);
            var listBindings = new List<ClassicListWebPartBindingSnapshot>();
            foreach (var webPart in sourceCapture.WebParts)
            {
                if (!ClassicListWebPartBindingParser.IsListBound(webPart))
                {
                    continue;
                }

                var binding = ClassicListWebPartBindingParser.Parse(
                    webPart,
                    sourceCapture.Identity.WebId,
                    sourceCapture.Identity.WebUrl,
                    sourceCapture.Identity.PageServerRelativeUrl);
                foreach (var issue in binding.Issues)
                {
                    blockers.Add(issue.Code + ": " + issue.Message);
                }
                if (binding.Binding != null)
                {
                    listBindings.Add(binding.Binding);
                }
            }

            var listClosure = ListDependencyClosureSnapshotReader.Read(
                sourceContext,
                listBindings,
                options.MaximumDependencyBytes,
                artifactStore,
                blockers,
                warnings);
            SourceSiteCollectionSnapshot sourceTopology = null;
            try
            {
                sourceTopology = SourceTopologySnapshotReader.CaptureRequiredWebClosure(
                    sourceContext,
                    listClosure.RequiredSourceWebIds.Concat(new[] { sourceCapture.Identity.WebId }));
            }
            catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
            {
                blockers.Add("Source topology closure could not be captured: " + exception.Message);
            }

            var references = PageReferenceSnapshotReader.Read(
                sourceContext,
                sourceCapture.Identity,
                sourceTopology,
                sourceCapture.PublishingPageContent,
                sourceCapture.WebParts,
                options,
                warnings);
            var afterFence = SourcePageFenceReader.Read(sourceContext, sourcePagePath);
            if (!SourcePageFenceReader.Equals(sourceCapture.SourceFence, afterFence))
            {
                blockers.Add("The source page changed while it was being exported. Discard this snapshot and export again.");
            }

            var runtime = PageRuntimeResolver.Resolve(
                sourceCapture.PageArtifact,
                sourceCapture.Layout?.PageDirective,
                sourceCapture.Identity.ContentTypeId);
            foreach (var diagnostic in runtime.Diagnostics)
            {
                warnings.Add(diagnostic);
            }
            var snapshot = new PublishingPageCaptureBundle
            {
                CapturePolicy = new PageCaptureOptions
                {
                    SourcePageServerRelativeUrl = sourcePagePath,
                    IncludeWebParts = options.IncludeWebParts,
                    MaximumDependencyBytes = options.MaximumDependencyBytes
                },
                Source = sourceCapture.Identity,
                PageArtifact = sourceCapture.PageArtifact,
                Runtime = runtime,
                ProfileSignals = PublishingPageProfileSignalProjector.Project(
                    sourceCapture.Identity,
                    sourceCapture.Layout,
                    sourceCapture.Fields),
                Layout = sourceCapture.Layout,
                PublishingPageContent = sourceCapture.PublishingPageContent,
                PublishingPageContentSha256 = PublishingPageDigest.ComputeSha256(sourceCapture.PublishingPageContent ?? string.Empty),
                Fields = sourceCapture.Fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToList(),
                WebParts = sourceCapture.WebParts
                    .OrderBy(webPart => webPart.ZoneId, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(webPart => webPart.ZoneIndex)
                    .ThenBy(webPart => webPart.Id)
                    .ToList(),
                ListWebPartBindings = listBindings
                    .OrderBy(binding => binding.SourceWebPartId)
                    .ToList(),
                ListDependencies = listClosure.Dependencies,
                ListLookupDependencies = listClosure.LookupDependencies,
                SourceTopology = sourceTopology,
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
            snapshot.IngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);

            var selection = workflowPolicy.Select(sourceCapture.Identity.ContentTypeId);
            return new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Selection = selection,
                SelectionDigest = PublishingPageDigest.ComputeSelectionDigest(selection),
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };
        }

        private static void ValidateOptions(PageCaptureOptions options)
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
