using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutSnapshotReader
    {
        public static PublishingPageLayoutSnapshot Read(
            ClientContext context,
            string layoutUrl,
            string description,
            IMigrationArtifactStore artifactStore,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            if (string.IsNullOrWhiteSpace(layoutUrl))
            {
                return Missing(layoutUrl, description, "PublishingPageLayout is unavailable on the source page.", blockers);
            }

            var rootWeb = context.Site.RootWeb;
            context.Load(context.Web, value => value.Url);
            context.Load(rootWeb, value => value.Url, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var serverRelativeUrl = ResolveServerRelativeUrl(rootWeb.Url, layoutUrl);
            Uri absoluteLayoutUrl;
            var rootUri = new Uri(rootWeb.Url);
            if (Uri.TryCreate(layoutUrl, UriKind.Absolute, out absoluteLayoutUrl)
                && (!string.Equals(absoluteLayoutUrl.Scheme, rootUri.Scheme, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(absoluteLayoutUrl.Host, rootUri.Host, StringComparison.OrdinalIgnoreCase)
                    || absoluteLayoutUrl.Port != rootUri.Port))
            {
                return Failure(
                    layoutUrl,
                    serverRelativeUrl,
                    description,
                    PublishingPageLayoutEvidenceState.Failed,
                    "PublishingPageLayout points outside the source site collection origin.",
                    blockers);
            }

            var galleryRoot = rootWeb.ServerRelativeUrl == "/"
                ? "/_catalogs/masterpage/"
                : rootWeb.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/masterpage/";
            if (!serverRelativeUrl.StartsWith(galleryRoot, StringComparison.OrdinalIgnoreCase))
            {
                return Failure(
                    layoutUrl,
                    serverRelativeUrl,
                    description,
                    PublishingPageLayoutEvidenceState.Failed,
                    $"PublishingPageLayout is outside the source master page gallery '{galleryRoot}'.",
                    blockers);
            }

            try
            {
                var file = rootWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
                var item = file.ListItemAllFields;
                var stream = file.OpenBinaryStream();
                context.Load(file,
                    value => value.Exists,
                    value => value.UniqueId,
                    value => value.Name,
                    value => value.ServerRelativeUrl,
                    value => value.UIVersionLabel,
                    value => value.Length,
                    value => value.CustomizedPageStatus,
                    value => value.CheckOutType,
                    value => value.Level);
                context.Load(item);
                context.ExecuteQueryRetry();
                if (!file.Exists || stream.Value == null)
                {
                    return Missing(layoutUrl, description, "The source Page Layout file or its exact bytes are unavailable.", blockers);
                }

                byte[] bytes;
                using (stream.Value)
                using (var buffer = new MemoryStream())
                {
                    stream.Value.CopyTo(buffer);
                    bytes = buffer.ToArray();
                }

                var markup = PublishingPageLayoutMarkupParser.Parse(PublishingPageLayoutEncoding.Decode(bytes));
                var association = ParseAssociation(FieldString(item, "PublishingAssociatedContentType"));
                var artifact = artifactStore == null
                    ? MigrationArtifact.Describe(bytes, "application/vnd.ms-aspx", file.Name)
                    : Put(artifactStore, bytes, file.Name);
                var schemaDiagnostics = new List<string>();
                var schema = association == null
                    ? null
                    : ContentTypeSchemaSnapshotReader.Read(context, rootWeb, association.Value.Id, markup.RequiredFieldNames, schemaDiagnostics);
                foreach (var diagnostic in schemaDiagnostics)
                {
                    warnings?.Add(diagnostic);
                }

                var sourceWebUrl = new Uri(context.Web.Url);
                var sourceSiteCollectionUrl = new Uri(rootWeb.Url);
                var resources = markup.ResourceReferences
                    .Select(reference => PublishingPageLayoutResourceSnapshotReader.Read(
                        context,
                        sourceWebUrl,
                        sourceSiteCollectionUrl,
                        reference,
                        artifactStore))
                    .ToList();
                return new PublishingPageLayoutSnapshot
                {
                    EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                    Availability = EvidenceAvailability.Captured,
                    Url = layoutUrl,
                    ServerRelativeUrl = file.ServerRelativeUrl,
                    Description = description,
                    FileUniqueId = file.UniqueId,
                    CustomizedPageStatus = (int)file.CustomizedPageStatus,
                    SetupPath = FieldString(item, "SetupPath"),
                    FileName = file.Name,
                    ItemContentTypeId = FieldString(item, "ContentTypeId"),
                    AssociatedContentTypeName = association?.Name,
                    AssociatedContentTypeId = association?.Id,
                    Title = FieldString(item, "Title"),
                    Level = file.Level.ToString(),
                    CheckOutType = file.CheckOutType.ToString(),
                    VersionLabel = file.UIVersionLabel,
                    Bytes = artifact,
                    ContentBase64 = artifactStore == null ? Convert.ToBase64String(bytes) : null,
                    PageDirective = markup.PageDirective,
                    Registrations = markup.Registrations,
                    Controls = markup.Controls,
                    Zones = markup.Zones,
                    ResourceReferences = markup.ResourceReferences,
                    ResourceArtifacts = resources,
                    AssociatedContentTypeSchema = schema,
                    Diagnostics = schemaDiagnostics
                };
            }
            catch (ServerException exception)
            {
                var accessDenied = exception.ServerErrorCode == -2147024891
                    || exception.Message.IndexOf("Access denied", StringComparison.OrdinalIgnoreCase) >= 0;
                var missing = exception.ServerErrorCode == -2147024894
                    || string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal);
                var state = accessDenied ? PublishingPageLayoutEvidenceState.AccessDenied
                    : missing ? PublishingPageLayoutEvidenceState.Missing
                    : PublishingPageLayoutEvidenceState.Failed;
                return Failure(layoutUrl, serverRelativeUrl, description, state, exception.Message, blockers);
            }
        }

        private static ArtifactReference Put(IMigrationArtifactStore store, byte[] bytes, string name)
        {
            using (var content = new MemoryStream(bytes, false))
            {
                return store.Put(content, "application/vnd.ms-aspx", name);
            }
        }

        private static string ResolveServerRelativeUrl(string rootWebUrl, string layoutUrl)
        {
            Uri absolute;
            if (Uri.TryCreate(layoutUrl, UriKind.Absolute, out absolute))
            {
                return Uri.UnescapeDataString(absolute.AbsolutePath);
            }

            if (layoutUrl.StartsWith("/", StringComparison.Ordinal))
            {
                return Uri.UnescapeDataString(layoutUrl.Split('?', '#')[0]);
            }

            return Uri.UnescapeDataString(new Uri(new Uri(rootWebUrl.TrimEnd('/') + "/"), layoutUrl).AbsolutePath);
        }

        private static (string Name, string Id)? ParseAssociation(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var parts = value.Split(new[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
            for (var index = 0; index + 1 < parts.Length; index++)
            {
                if (parts[index + 1].Trim().StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    return (parts[index].Trim(), parts[index + 1].Trim());
                }
            }

            var comma = value.LastIndexOf(", 0x", StringComparison.OrdinalIgnoreCase);
            return comma <= 0 ? ((string, string)?)null : (value.Substring(0, comma).Trim(), value.Substring(comma + 2).Trim());
        }

        private static string FieldString(ListItem item, string name)
        {
            object value;
            return item.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }

        private static PublishingPageLayoutSnapshot Missing(
            string url,
            string description,
            string diagnostic,
            ICollection<string> blockers)
        {
            blockers?.Add(diagnostic);
            return Failure(url, null, description, PublishingPageLayoutEvidenceState.Missing, diagnostic, null);
        }

        private static PublishingPageLayoutSnapshot Failure(
            string url,
            string serverRelativeUrl,
            string description,
            PublishingPageLayoutEvidenceState state,
            string diagnostic,
            ICollection<string> blockers)
        {
            blockers?.Add(diagnostic);
            return new PublishingPageLayoutSnapshot
            {
                Url = url,
                ServerRelativeUrl = serverRelativeUrl,
                Description = description,
                EvidenceState = state,
                Availability = state == PublishingPageLayoutEvidenceState.MetadataOnly
                    ? EvidenceAvailability.Partial
                    : EvidenceAvailability.Unavailable,
                Diagnostics = new List<string> { diagnostic }
            };
        }
    }
}
