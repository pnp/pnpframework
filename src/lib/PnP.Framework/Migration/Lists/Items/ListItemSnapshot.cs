using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Items.Protection;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Lists.Items
{
    public enum ListDocumentObjectKind
    {
        File = 1,
        Folder = 2
    }

    /// <summary>
    /// Describes what SharePoint returned for a captured binary request. The
    /// artifact digest always seals those exact returned bytes; this value tells
    /// planning whether those bytes are also a stable representation of the
    /// logical source file.
    /// </summary>
    public enum ListBinaryRepresentationKind
    {
        Unclassified = 0,
        OrdinaryFilePayload = 1,
        InformationRightsManagedEnvelope = 2
    }

    public sealed class ListBinaryContentIdentitySnapshot
    {
        public string QuickXorHash { get; set; }

        public string ContentTag { get; set; }

        public string EvidenceSource { get; set; }
    }

    public sealed class ListBinaryArtifactSnapshot
    {
        public ArtifactReference Artifact { get; set; }

        public string ContentBase64 { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public ListBinaryRepresentationKind RepresentationKind { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public ListBinaryContentIdentitySnapshot LogicalContentIdentity { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public IList<LiteralHttpArchivedContentEvidence> ArchivedContentEvidence { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class ListAttachmentSnapshot
    {
        public string FileName { get; set; }

        public string ServerRelativeUrl { get; set; }

        public ListBinaryArtifactSnapshot Content { get; set; }
    }

    public sealed class ListDocumentSnapshot
    {
        public ListDocumentObjectKind Kind { get; set; }

        public string Name { get; set; }

        public string ServerRelativeUrl { get; set; }

        public long Length { get; set; }

        public int MajorVersion { get; set; }

        public int MinorVersion { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public ListDocumentInformationProtectionSnapshot InformationProtection { get; set; }

        public ListBinaryArtifactSnapshot Content { get; set; }
    }

    public sealed class ListItemSnapshot
    {
        public int SourceItemId { get; set; }

        public Guid? SourceUniqueId { get; set; }

        public IList<ListItemValueSnapshot> Values { get; set; } = new List<ListItemValueSnapshot>();

        public IList<ListAttachmentSnapshot> Attachments { get; set; } = new List<ListAttachmentSnapshot>();

        public ListDocumentSnapshot Document { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
