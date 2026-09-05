using PnP.Framework.Migration.Evidence;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Items.Protection
{
    /// <summary>
    /// A fidelity-first projection of the document-level Microsoft Information
    /// Protection assignment retained in SharePoint item metadata. The original
    /// field values remain in the List item snapshot; these properties expose the
    /// external policy relationship as an independently planned ingredient.
    /// </summary>
    public sealed class ListDocumentInformationProtectionSnapshot
    {
        public string LabelId { get; set; }

        public string AssignmentMethod { get; set; }

        public string HasUserDefinedProtection { get; set; }

        public string OwnerEmail { get; set; }

        public string LabelHash { get; set; }

        public string PromotionCtagVersion { get; set; }

        public string DecryptSkipReason { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
