using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Views
{
    internal static class ListViewRenderingResourceSnapshotReader
    {
        public static IList<ListViewRenderingResourceSnapshot> Read(
            ClientContext context,
            Web sourceWeb,
            Web sourceRootWeb,
            IEnumerable<ListViewSnapshot> views,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ICollection<string> warnings)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (sourceWeb == null)
            {
                throw new ArgumentNullException(nameof(sourceWeb));
            }
            if (sourceRootWeb == null)
            {
                throw new ArgumentNullException(nameof(sourceRootWeb));
            }
            if (maximumBytes <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }

            var sourceWebUri = new Uri(sourceWeb.Url.TrimEnd('/') + "/");
            var sourceSiteUri = new Uri(sourceRootWeb.Url.TrimEnd('/') + "/");
            var resources = new Dictionary<string, ListViewRenderingResourceSnapshot>(StringComparer.Ordinal);
            foreach (var view in (views ?? Array.Empty<ListViewSnapshot>()).Where(value => value != null))
            {
                var bindings = Extract(view.JsLink, "JSLink", ListViewRenderingResourceKind.JavaScript)
                    .Concat(Extract(view.XslLink, "XslLink", ListViewRenderingResourceKind.Xsl))
                    .ToArray();
                foreach (var binding in bindings)
                {
                    var sourceUri = ResolveSourceUri(sourceWebUri, sourceSiteUri, binding.OriginalReference);
                    var resourceId = ResourceId(sourceUri, binding.OriginalReference);
                    view.RenderingResourceBindings.Add(new ListViewRenderingResourceBindingSnapshot
                    {
                        SourceProperty = binding.SourceProperty,
                        OriginalReference = binding.OriginalReference,
                        ResourceId = resourceId
                    });
                    if (!resources.ContainsKey(resourceId))
                    {
                        resources[resourceId] = Capture(
                            context,
                            sourceWeb,
                            sourceRootWeb,
                            sourceWebUri,
                            sourceSiteUri,
                            sourceUri,
                            resourceId,
                            binding.Kind,
                            maximumBytes,
                            artifactStore,
                            warnings);
                    }
                }
                view.RenderingResourceBindings = view.RenderingResourceBindings
                    .GroupBy(value => value.SourceProperty + "\u001f" + value.OriginalReference, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .OrderBy(value => value.SourceProperty, StringComparer.Ordinal)
                    .ThenBy(value => value.OriginalReference, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            return resources.Values
                .OrderBy(value => value.Id, StringComparer.Ordinal)
                .ToList();
        }

        internal static bool IsCustomReference(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && (value.IndexOf('/') >= 0
                    || value.IndexOf('\\') >= 0
                    || value.StartsWith("~", StringComparison.Ordinal));
        }

        internal static Uri ResolveSourceUri(Uri sourceWebUri, Uri sourceSiteUri, string reference)
        {
            if (sourceWebUri == null || sourceSiteUri == null || string.IsNullOrWhiteSpace(reference))
            {
                return null;
            }

            var value = reference.Trim().Replace('\\', '/');
            if (value.StartsWith("~sitecollection/", StringComparison.OrdinalIgnoreCase))
            {
                return new Uri(sourceSiteUri, value.Substring("~sitecollection/".Length));
            }
            if (value.StartsWith("~site/", StringComparison.OrdinalIgnoreCase))
            {
                return new Uri(sourceWebUri, value.Substring("~site/".Length));
            }
            if (Uri.TryCreate(value, UriKind.Absolute, out var absolute))
            {
                return absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps
                    ? absolute
                    : null;
            }
            if (value.StartsWith("/", StringComparison.Ordinal))
            {
                return new Uri(new Uri(sourceWebUri.GetLeftPart(UriPartial.Authority) + "/"), value.TrimStart('/'));
            }
            return IsCustomReference(value) ? new Uri(sourceWebUri, value) : null;
        }

        private static ListViewRenderingResourceSnapshot Capture(
            ClientContext context,
            Web sourceWeb,
            Web sourceRootWeb,
            Uri sourceWebUri,
            Uri sourceSiteUri,
            Uri sourceUri,
            string resourceId,
            ListViewRenderingResourceKind kind,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ICollection<string> warnings)
        {
            var snapshot = new ListViewRenderingResourceSnapshot
            {
                Id = resourceId,
                Kind = kind,
                SourceAbsoluteUrl = sourceUri?.AbsoluteUri,
                Availability = EvidenceAvailability.Unavailable
            };
            if (sourceUri == null)
            {
                snapshot.Diagnostics.Add("The custom View rendering reference could not be resolved to an HTTP(S) source URI.");
                return snapshot;
            }
            if (!string.Equals(sourceUri.Host, sourceWebUri.Host, StringComparison.OrdinalIgnoreCase))
            {
                snapshot.Diagnostics.Add("External View rendering resources require a separately reviewed capture provider.");
                return snapshot;
            }

            var sourcePath = Uri.UnescapeDataString(sourceUri.AbsolutePath);
            var sourceSitePath = Uri.UnescapeDataString(sourceSiteUri.AbsolutePath).TrimEnd('/');
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            snapshot.SourceServerRelativeUrl = sourcePath;
            if (!IsWithin(sourcePath, sourceSitePath))
            {
                snapshot.Diagnostics.Add("The custom View rendering resource is outside the source Site Collection.");
                return snapshot;
            }

            try
            {
                ListBinaryArtifactSnapshot binary;
                if (!IsWithin(sourcePath, sourceWebPath) && sourceWeb.Id != sourceRootWeb.Id)
                {
                    using (var ownerContext = context.Clone(sourceRootWeb.Url))
                    {
                        binary = ReadBinary(ownerContext, ownerContext.Web, sourcePath, maximumBytes, artifactStore, kind);
                    }
                }
                else
                {
                    binary = ReadBinary(context, sourceWeb, sourcePath, maximumBytes, artifactStore, kind);
                }
                snapshot.Artifact = binary.Artifact;
                snapshot.ContentBase64 = binary.ContentBase64;
                snapshot.Availability = binary.Availability;
                foreach (var diagnostic in binary.Diagnostics)
                {
                    snapshot.Diagnostics.Add(diagnostic);
                }
            }
            catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
            {
                snapshot.Diagnostics.Add(exception.Message);
                warnings?.Add($"View rendering resource '{sourceUri}' could not be captured: {exception.Message}");
            }
            return snapshot;
        }

        private static ListBinaryArtifactSnapshot ReadBinary(
            ClientContext context,
            Web owner,
            string sourcePath,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ListViewRenderingResourceKind kind)
        {
            var file = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(sourcePath));
            return ListBinaryArtifactReader.Read(
                context,
                file,
                maximumBytes,
                artifactStore,
                MediaType(kind),
                Path.GetFileName(sourcePath),
                sourcePath);
        }

        private static IEnumerable<RenderingBindingCandidate> Extract(
            string composite,
            string sourceProperty,
            ListViewRenderingResourceKind kind)
        {
            return (composite ?? string.Empty)
                .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim())
                .Where(IsCustomReference)
                .Select(value => new RenderingBindingCandidate
                {
                    SourceProperty = sourceProperty,
                    OriginalReference = value,
                    Kind = kind
                });
        }

        private static string ResourceId(Uri sourceUri, string originalReference)
        {
            var identity = sourceUri == null
                ? "unresolved:" + (originalReference ?? string.Empty)
                : sourceUri.GetLeftPart(UriPartial.Path);
            return MigrationDigest.ComputeSha256(identity.ToLowerInvariant());
        }

        private static bool IsWithin(string value, string parent)
        {
            var normalizedParent = string.IsNullOrEmpty(parent) ? string.Empty : parent.TrimEnd('/');
            return string.IsNullOrEmpty(normalizedParent)
                || string.Equals(value, normalizedParent, StringComparison.OrdinalIgnoreCase)
                || value.StartsWith(normalizedParent + "/", StringComparison.OrdinalIgnoreCase);
        }

        private static string MediaType(ListViewRenderingResourceKind kind)
        {
            return kind == ListViewRenderingResourceKind.JavaScript
                ? "application/javascript"
                : kind == ListViewRenderingResourceKind.Xsl
                    ? "application/xslt+xml"
                    : kind == ListViewRenderingResourceKind.StyleSheet
                        ? "text/css"
                        : "application/octet-stream";
        }

        private sealed class RenderingBindingCandidate
        {
            public string SourceProperty { get; set; }

            public string OriginalReference { get; set; }

            public ListViewRenderingResourceKind Kind { get; set; }
        }
    }
}
