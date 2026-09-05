using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

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

            var pageRootWeb = context.Site.RootWeb;
            context.Load(context.Web, value => value.Url);
            context.Load(pageRootWeb, value => value.Url, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            PublishingPageLayoutSourceLocation location;
            string locationDiagnostic;
            if (!PublishingPageLayoutSourceLocation.TryResolve(
                new Uri(pageRootWeb.Url),
                layoutUrl,
                out location,
                out locationDiagnostic))
            {
                return Failure(
                    layoutUrl,
                    null,
                    description,
                    PublishingPageLayoutEvidenceState.Failed,
                    locationDiagnostic,
                    blockers);
            }

            var pageWebUrl = new Uri(context.Web.Url);
            var pageSiteCollectionUrl = new Uri(pageRootWeb.Url);
            if (!location.IsExternalToPageSiteCollection)
            {
                return ReadFromOwner(
                    context,
                    context,
                    pageWebUrl,
                    pageSiteCollectionUrl,
                    layoutUrl,
                    description,
                    location,
                    artifactStore,
                    blockers,
                    warnings);
            }

            using (var ownerContext = context.Clone(location.OwnerSiteCollectionUrl))
            {
                return ReadFromOwner(
                    ownerContext,
                    context,
                    pageWebUrl,
                    pageSiteCollectionUrl,
                    layoutUrl,
                    description,
                    location,
                    artifactStore,
                    blockers,
                    warnings);
            }
        }

        private static PublishingPageLayoutSnapshot ReadFromOwner(
            ClientContext ownerContext,
            ClientContext pageContext,
            Uri pageWebUrl,
            Uri pageSiteCollectionUrl,
            string layoutUrl,
            string description,
            PublishingPageLayoutSourceLocation location,
            IMigrationArtifactStore artifactStore,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            try
            {
                var rootWeb = ownerContext.Site.RootWeb;
                ownerContext.Load(rootWeb, value => value.Url, value => value.ServerRelativeUrl);
                ownerContext.ExecuteQueryRetry();
                var galleryRoot = rootWeb.ServerRelativeUrl == "/"
                    ? "/_catalogs/masterpage/"
                    : rootWeb.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/masterpage/";
                if (!location.ServerRelativeUrl.StartsWith(galleryRoot, StringComparison.OrdinalIgnoreCase))
                {
                    return Failure(
                        layoutUrl,
                        location.ServerRelativeUrl,
                        description,
                        PublishingPageLayoutEvidenceState.Failed,
                        $"PublishingPageLayout owner resolved to '{rootWeb.Url}', but the file is outside its master page gallery '{galleryRoot}'.",
                        blockers,
                        location);
                }

                var file = rootWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(location.ServerRelativeUrl));
                var item = file.ListItemAllFields;
                var stream = file.OpenBinaryStream();
                ownerContext.Load(file,
                    value => value.Exists,
                    value => value.UniqueId,
                    value => value.Name,
                    value => value.ServerRelativeUrl,
                    value => value.UIVersionLabel,
                    value => value.Length,
                    value => value.CustomizedPageStatus,
                    value => value.CheckOutType,
                    value => value.Level);
                ownerContext.Load(item);
                ownerContext.ExecuteQueryRetry();
                if (!file.Exists || stream.Value == null)
                {
                    return Missing(layoutUrl, description, "The source Page Layout file or its exact bytes are unavailable.", blockers, location);
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
                    : ContentTypeSchemaSnapshotReader.Read(ownerContext, rootWeb, association.Value.Id, markup.RequiredFieldIdentifiers, schemaDiagnostics);
                foreach (var diagnostic in schemaDiagnostics)
                {
                    warnings?.Add(diagnostic);
                }

                var resources = markup.ResourceReferences
                    .Select(reference => PublishingPageLayoutResourceSnapshotReader.Read(
                        pageContext,
                        pageWebUrl,
                        pageSiteCollectionUrl,
                        reference,
                        artifactStore))
                    .ToList();
                return new PublishingPageLayoutSnapshot
                {
                    EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                    Availability = EvidenceAvailability.Captured,
                    Url = layoutUrl,
                    ServerRelativeUrl = file.ServerRelativeUrl,
                    OwnerSiteCollectionUrl = rootWeb.Url,
                    ExternalToPageSiteCollection = location.IsExternalToPageSiteCollection,
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
            catch (WebException exception)
            {
                using (var response = exception.Response as HttpWebResponse)
                {
                    var statusCode = response == null ? 0 : (int)response.StatusCode;
                    if (statusCode != 401 && statusCode != 403)
                    {
                        throw;
                    }

                    var requestUri = response.ResponseUri?.AbsoluteUri
                        ?? location.AbsoluteLayoutUrl.AbsoluteUri;
                    var evidence = LiteralHttpAuthorizationEvidence.Create(
                        "capture-page-layout-owner",
                        requestUri,
                        statusCode,
                        DateTimeOffset.UtcNow);
                    return AuthorizationFailure(
                        layoutUrl,
                        description,
                        location,
                        evidence,
                        blockers);
                }
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
                return Failure(layoutUrl, location.ServerRelativeUrl, description, state, exception.Message, blockers, location);
            }
        }

        private static ArtifactReference Put(IMigrationArtifactStore store, byte[] bytes, string name)
        {
            using (var content = new MemoryStream(bytes, false))
            {
                return store.Put(content, "application/vnd.ms-aspx", name);
            }
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
            ICollection<string> blockers,
            PublishingPageLayoutSourceLocation location = null)
        {
            blockers?.Add(diagnostic);
            return Failure(
                url,
                location?.ServerRelativeUrl,
                description,
                PublishingPageLayoutEvidenceState.Missing,
                diagnostic,
                null,
                location);
        }

        private static PublishingPageLayoutSnapshot AuthorizationFailure(
            string url,
            string description,
            PublishingPageLayoutSourceLocation location,
            LiteralHttpAuthorizationEvidence evidence,
            ICollection<string> blockers)
        {
            var diagnostic = $"Page Layout owner request returned literal HTTP {evidence.HttpStatusCode}.";
            blockers?.Add(diagnostic);
            return new PublishingPageLayoutSnapshot
            {
                Url = url,
                ServerRelativeUrl = location.ServerRelativeUrl,
                OwnerSiteCollectionUrl = location.OwnerSiteCollectionUrl.AbsoluteUri.TrimEnd('/'),
                ExternalToPageSiteCollection = location.IsExternalToPageSiteCollection,
                Description = description,
                EvidenceState = PublishingPageLayoutEvidenceState.AuthorizationBlocked,
                Availability = EvidenceAvailability.Unavailable,
                AuthorizationEvidence = evidence,
                Diagnostics = new List<string> { diagnostic }
            };
        }

        private static PublishingPageLayoutSnapshot Failure(
            string url,
            string serverRelativeUrl,
            string description,
            PublishingPageLayoutEvidenceState state,
            string diagnostic,
            ICollection<string> blockers,
            PublishingPageLayoutSourceLocation location = null)
        {
            blockers?.Add(diagnostic);
            return new PublishingPageLayoutSnapshot
            {
                Url = url,
                ServerRelativeUrl = serverRelativeUrl,
                OwnerSiteCollectionUrl = location?.OwnerSiteCollectionUrl.AbsoluteUri.TrimEnd('/'),
                ExternalToPageSiteCollection = location?.IsExternalToPageSiteCollection == true,
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
