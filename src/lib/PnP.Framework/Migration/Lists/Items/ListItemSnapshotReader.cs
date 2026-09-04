using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Items.Protection;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListItemSnapshotReader
    {
        public static IList<ListItemSnapshot> Read(
            ClientContext context,
            List list,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ICollection<string> warnings)
        {
            var result = new List<ListItemSnapshot>();
            ListItemCollectionPosition position = null;
            do
            {
                var page = list.GetItems(new CamlQuery
                {
                    ViewXml = "<View Scope='RecursiveAll'><RowLimit Paged='TRUE'>5000</RowLimit></View>",
                    ListItemCollectionPosition = position
                });
                context.Load(page);
                context.ExecuteQueryRetry();
                if (list.BaseType == BaseType.DocumentLibrary)
                {
                    foreach (var item in page)
                    {
                        if (item.FileSystemObjectType == FileSystemObjectType.File)
                        {
                            context.Load(item.File, value => value.Name, value => value.ServerRelativeUrl, value => value.Length, value => value.MajorVersion, value => value.MinorVersion);
                        }
                        else
                        {
                            context.Load(item.Folder, value => value.Name, value => value.ServerRelativeUrl);
                        }
                    }
                    context.ExecuteQueryRetry();
                }
                if (list.EnableAttachments)
                {
                    foreach (var item in page.Where(HasAttachments))
                    {
                        context.Load(item.AttachmentFiles, values => values.Include(value => value.FileName, value => value.ServerRelativeUrl));
                    }
                    context.ExecuteQueryRetry();
                }

                foreach (var item in page)
                {
                    var snapshot = new ListItemSnapshot
                    {
                        SourceItemId = item.Id,
                        SourceUniqueId = ReadUniqueId(item),
                        Values = item.FieldValues.OrderBy(value => value.Key, StringComparer.OrdinalIgnoreCase)
                            .Select(value => ListItemValueSerializer.Serialize(value.Key, value.Value)).ToList(),
                        Attachments = CaptureAttachments(context, item, maximumBytes, artifactStore),
                        Document = list.BaseType == BaseType.DocumentLibrary
                            ? CaptureDocument(context, item, maximumBytes, artifactStore)
                            : null
                    };
                    var unavailableBinary = snapshot.Attachments.Any(value => value.Content == null || value.Content.Availability != EvidenceAvailability.Captured)
                        || (snapshot.Document != null && snapshot.Document.Kind == ListDocumentObjectKind.File
                            && (snapshot.Document.Content == null || snapshot.Document.Content.Availability != EvidenceAvailability.Captured));
                    if (unavailableBinary || snapshot.Values.Any(value => value.Availability != EvidenceAvailability.Captured))
                    {
                        snapshot.Availability = EvidenceAvailability.Partial;
                    }
                    if (unavailableBinary)
                    {
                        warnings.Add("List item " + item.Id + " has document or attachment bytes that could not be captured exactly.");
                    }
                    if (snapshot.Document?.Content?.RepresentationKind
                        == ListBinaryRepresentationKind.InformationRightsManagedEnvelope)
                    {
                        warnings.Add("List item " + item.Id + " returned an Information Rights Management envelope. The exact response bytes and stable logical-content identity are retained, but cross-site replay and verification remain pending.");
                    }
                    result.Add(snapshot);
                }
                position = page.ListItemCollectionPosition;
            }
            while (position != null);

            return result.OrderBy(value => value.SourceItemId).ToList();
        }

        private static IList<ListAttachmentSnapshot> CaptureAttachments(ClientContext context, ListItem item, long maximumBytes, IMigrationArtifactStore artifactStore)
        {
            if (!HasAttachments(item))
            {
                return new List<ListAttachmentSnapshot>();
            }
            return item.AttachmentFiles.AsEnumerable().Select(attachment =>
            {
                var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(Uri.UnescapeDataString(attachment.ServerRelativeUrl)));
                return new ListAttachmentSnapshot
                {
                    FileName = attachment.FileName,
                    ServerRelativeUrl = attachment.ServerRelativeUrl,
                    Content = ListBinaryArtifactReader.Read(context, file, maximumBytes, artifactStore, "application/octet-stream", attachment.FileName)
                };
            }).ToList();
        }

        private static ListDocumentSnapshot CaptureDocument(ClientContext context, ListItem item, long maximumBytes, IMigrationArtifactStore artifactStore)
        {
            if (item.FileSystemObjectType == FileSystemObjectType.Folder)
            {
                return new ListDocumentSnapshot
                {
                    Kind = ListDocumentObjectKind.Folder,
                    Name = item.Folder.Name,
                    ServerRelativeUrl = item.Folder.ServerRelativeUrl
                };
            }
            var content = ListBinaryArtifactReader.Read(
                context,
                item.File,
                maximumBytes,
                artifactStore,
                ListBinaryArtifactReader.MediaType(item.File.Name),
                item.File.Name,
                item.File.ServerRelativeUrl,
                ReadArchiveStatus(item));
            if (content?.RepresentationKind == ListBinaryRepresentationKind.InformationRightsManagedEnvelope)
            {
                content.LogicalContentIdentity = ListBinaryContentIdentityReader.Read(item.FieldValues);
                if (content.LogicalContentIdentity == null)
                {
                    content.Diagnostics.Add("RightsManagedLogicalContentIdentityUnavailable: SharePoint returned a DRM envelope, but MetaInfo contained no cTag or QuickXorHash.");
                }
            }
            if (content?.Artifact != null && item.File.Length != content.Artifact.Length)
            {
                if (content.RepresentationKind == ListBinaryRepresentationKind.InformationRightsManagedEnvelope)
                {
                    content.Diagnostics.Add("RightsManagedEnvelopeLengthMismatch: logicalFileLength=" + item.File.Length
                        + "; returnedEnvelopeLength=" + content.Artifact.Length + ".");
                }
                else
                {
                    content.Availability = EvidenceAvailability.Partial;
                    content.Diagnostics.Add("DocumentMetadataLengthMismatch: metadataLength=" + item.File.Length
                        + "; payloadLength=" + content.Artifact.Length + ".");
                }
            }
            return new ListDocumentSnapshot
            {
                Kind = ListDocumentObjectKind.File,
                Name = item.File.Name,
                ServerRelativeUrl = item.File.ServerRelativeUrl,
                Length = item.File.Length,
                MajorVersion = item.File.MajorVersion,
                MinorVersion = item.File.MinorVersion,
                InformationProtection = ListDocumentInformationProtectionSnapshotReader.Read(item.FieldValues),
                Content = content
            };
        }

        private static bool HasAttachments(ListItem item)
        {
            object value;
            return item.FieldValues.TryGetValue("Attachments", out value) && value is bool && (bool)value;
        }

        private static string ReadArchiveStatus(ListItem item)
        {
            foreach (var value in item.FieldValues)
            {
                if (string.Equals(
                    value.Key,
                    "_FileArchiveStatus",
                    StringComparison.OrdinalIgnoreCase))
                {
                    return Convert.ToString(value.Value);
                }
            }
            return null;
        }

        private static Guid? ReadUniqueId(ListItem item)
        {
            object value;
            if (!item.FieldValues.TryGetValue("GUID", out value) || value == null)
            {
                return null;
            }
            if (value is Guid)
            {
                return (Guid)value;
            }
            Guid parsed;
            return Guid.TryParse(Convert.ToString(value), out parsed) ? parsed : (Guid?)null;
        }
    }
}
