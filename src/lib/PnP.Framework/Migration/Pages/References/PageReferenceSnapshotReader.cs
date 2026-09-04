using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.Packaging;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PnP.Framework.Migration.Pages.References
{
    internal static class PageReferenceSnapshotReader
    {
        private static readonly Regex CssUrlPattern = new Regex(
            @"url\(\s*(?:['""](?<url>.*?)['""]|(?<url>[^)]*?))\s*\)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static List<PageReferenceSnapshot> Read(
            ClientContext sourceContext,
            PageIdentity source,
            SourceSiteCollectionSnapshot sourceTopology,
            string pageContent,
            IEnumerable<ClassicWebPartSnapshot> webParts,
            PageCaptureOptions options,
            ICollection<string> warnings)
        {
            var candidates = ExtractHtmlReferences(pageContent);
            foreach (var webPart in webParts)
            {
                candidates.AddRange(ExtractTextReferences(webPart.ExportXml, $"webpart:{webPart.Id}"));
            }

            var sourceWebUri = new Uri(UrlUtility.EnsureTrailingSlash(source.WebUrl));
            var sourcePageUri = new Uri(sourceWebUri.GetLeftPart(UriPartial.Authority) + PagePath.Encode(source.PageServerRelativeUrl));
            var result = new List<PageReferenceSnapshot>();
            foreach (var candidate in candidates
                         .GroupBy(item => $"{item.Consumer}\n{item.Value}", StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                if (!TryResolveUri(sourcePageUri, candidate.Value, out var absoluteUri))
                {
                    continue;
                }

                result.Add(Capture(
                    sourceContext,
                    source,
                    sourceTopology,
                    sourceWebUri,
                    candidate,
                    absoluteUri,
                    options,
                    warnings));
            }

            return result;
        }

        public static bool IsSharePointRuntimePath(string serverRelativeUrl)
        {
            return serverRelativeUrl.StartsWith("/_layouts/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.StartsWith("/_vti_bin/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.StartsWith("/_api/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.IndexOf("/_catalogs/masterpage/", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static PageReferenceSnapshot Capture(
            ClientContext sourceContext,
            PageIdentity source,
            SourceSiteCollectionSnapshot sourceTopology,
            Uri sourceWebUri,
            ReferenceCandidate candidate,
            Uri absoluteUri,
            PageCaptureOptions options,
            ICollection<string> warnings)
        {
            var reference = new PageReferenceSnapshot
            {
                Id = PageDigest.ComputeSha256($"{candidate.Consumer}\n{absoluteUri.AbsoluteUri}"),
                OriginalValue = candidate.Value,
                SourceAbsoluteUrl = absoluteUri.AbsoluteUri,
                Consumer = candidate.Consumer,
                Kind = candidate.Kind,
                IsRenderableResource = candidate.IsRenderableResource,
                CaptureStatus = PageCaptureStatus.Captured
            };
            if (!string.Equals(sourceWebUri.Host, absoluteUri.Host, StringComparison.OrdinalIgnoreCase))
            {
                return reference;
            }

            var sourcePath = Uri.UnescapeDataString(absoluteUri.AbsolutePath);
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            reference.SourceServerRelativeUrl = sourcePath;
            if (!candidate.IsRenderableResource || IsSharePointRuntimePath(sourcePath))
            {
                return reference;
            }

            if (candidate.Kind == PageReferenceKind.IFrame)
            {
                reference.CaptureStatus = PageCaptureStatus.CapturedWithLimitations;
                reference.Diagnostics.Add("Same-tenant iframe dependencies require a separately reviewed page/application profile during planning.");
                return reference;
            }

            var owner = (sourceTopology?.Webs ?? Array.Empty<SourceWebSnapshot>())
                .Where(web => web != null
                    && !string.IsNullOrWhiteSpace(web.ServerRelativeUrl)
                    && PagePath.IsWithin(sourcePath, web.ServerRelativeUrl))
                .OrderByDescending(web => web.ServerRelativeUrl.Length)
                .FirstOrDefault();
            if (owner == null && PagePath.IsWithin(sourcePath, sourceWebPath))
            {
                owner = new SourceWebSnapshot
                {
                    WebId = source.WebId,
                    ServerRelativeUrl = source.WebServerRelativeUrl,
                    WebUrl = source.WebUrl
                };
            }
            if (owner == null)
            {
                reference.CaptureStatus = PageCaptureStatus.CapturedWithLimitations;
                reference.Diagnostics.Add(
                    "The resource owner is outside the captured source Site/Web topology closure.");
                return reference;
            }

            try
            {
                var ownerWeb = owner.WebId == source.WebId
                    ? sourceContext.Web
                    : sourceContext.Site.OpenWebById(owner.WebId);
                var payload = ReadFile(ownerWeb, sourceContext, sourcePath, options.MaximumDependencyBytes);
                reference.ContentBase64 = Convert.ToBase64String(payload);
                reference.ContentLength = payload.LongLength;
                reference.ContentSha256 = PageDigest.ComputeSha256(payload);
            }
            catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
            {
                reference.CaptureStatus = PageCaptureStatus.Failed;
                reference.Diagnostics.Add(exception.Message);
                warnings.Add($"Resource '{absoluteUri}' could not be captured and may block a later plan: {exception.Message}");
            }

            return reference;
        }

        private static byte[] ReadFile(
            Web web,
            ClientContext context,
            string serverRelativeUrl,
            long maximumBytes)
        {
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            context.Load(file, value => value.Exists, value => value.Length);
            var stream = file.OpenBinaryStream();
            context.ExecuteQueryRetry();
            if (!file.Exists || stream.Value == null)
            {
                throw new FileNotFoundException("The referenced SharePoint file was not found.", serverRelativeUrl);
            }

            if (file.Length > maximumBytes)
            {
                throw new InvalidOperationException($"The dependency is {file.Length} bytes, above the configured {maximumBytes}-byte limit.");
            }

            using (stream.Value)
            using (var output = new MemoryStream())
            {
                stream.Value.CopyTo(output);
                if (output.Length > maximumBytes)
                {
                    throw new InvalidOperationException($"The dependency is above the configured {maximumBytes}-byte limit.");
                }

                return output.ToArray();
            }
        }

        private static List<ReferenceCandidate> ExtractHtmlReferences(string html)
        {
            var result = new List<ReferenceCandidate>();
            if (string.IsNullOrWhiteSpace(html))
            {
                return result;
            }

            var document = new HtmlParser().ParseDocument(html);
            foreach (var element in document.All)
            {
                AddAttributeReference(result, element, "href", GetKind(element, "href"));
                AddAttributeReference(result, element, "src", GetKind(element, "src"));
                AddAttributeReference(result, element, "poster", PageReferenceKind.Media);
                AddAttributeReference(result, element, "data", PageReferenceKind.Object);
                var style = element.GetAttribute("style");
                if (!string.IsNullOrWhiteSpace(style))
                {
                    foreach (Match match in CssUrlPattern.Matches(style))
                    {
                        result.Add(new ReferenceCandidate
                        {
                            Consumer = $"{element.LocalName}[style]",
                            Kind = PageReferenceKind.Image,
                            Value = match.Groups["url"].Value.Trim(),
                            IsRenderableResource = true
                        });
                    }
                }
            }

            return result;
        }

        private static IEnumerable<ReferenceCandidate> ExtractTextReferences(string text, string consumer)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return Array.Empty<ReferenceCandidate>();
            }

            return Regex.Matches(text, @"https?://[^\s'""<>]+|(?<quote>['""])(?<path>/[^'""<>\s]+)\k<quote>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Cast<Match>()
                .Select(match => new ReferenceCandidate
                {
                    Consumer = consumer,
                    Kind = PageReferenceKind.Unknown,
                    Value = (match.Groups["path"].Success ? match.Groups["path"].Value : match.Value).TrimEnd('.', ',', ';', ')'),
                    IsRenderableResource = false
                })
                .ToArray();
        }

        private static void AddAttributeReference(
            ICollection<ReferenceCandidate> result,
            IElement element,
            string attributeName,
            PageReferenceKind kind)
        {
            var value = element.GetAttribute(attributeName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            result.Add(new ReferenceCandidate
            {
                Consumer = $"{element.LocalName}[{attributeName}]",
                Kind = kind,
                Value = value.Trim(),
                IsRenderableResource = kind != PageReferenceKind.Anchor
                    && kind != PageReferenceKind.Unknown
            });
        }

        private static PageReferenceKind GetKind(IElement element, string attributeName)
        {
            switch (element.LocalName.ToLowerInvariant())
            {
                case "a":
                case "area":
                    return PageReferenceKind.Anchor;
                case "img":
                    return PageReferenceKind.Image;
                case "script":
                    return PageReferenceKind.Script;
                case "link":
                    return PageReferenceKind.StyleSheet;
                case "iframe":
                    return PageReferenceKind.IFrame;
                case "object":
                    return PageReferenceKind.Object;
                case "audio":
                case "source":
                case "video":
                    return PageReferenceKind.Media;
                default:
                    return attributeName == "href"
                        ? PageReferenceKind.Anchor
                        : PageReferenceKind.Unknown;
            }
        }

        private static bool TryResolveUri(Uri sourcePageUri, string value, out Uri result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(value)
                || value.StartsWith("#", StringComparison.Ordinal)
                || value.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("tel:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return Uri.TryCreate(sourcePageUri, value, out result)
                && (result.Scheme == Uri.UriSchemeHttps || result.Scheme == Uri.UriSchemeHttp);
        }

        private sealed class ReferenceCandidate
        {
            public string Value { get; set; }

            public string Consumer { get; set; }

            public PageReferenceKind Kind { get; set; }

            public bool IsRenderableResource { get; set; }
        }
    }
}
