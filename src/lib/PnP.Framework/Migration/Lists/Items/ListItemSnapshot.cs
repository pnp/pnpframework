using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Items
{
    public enum ListDocumentObjectKind
    {
        File = 1,
        Folder = 2
    }

    public sealed class ListBinaryArtifactSnapshot
    {
        public ArtifactReference Artifact { get; set; }

        public string ContentBase64 { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

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
