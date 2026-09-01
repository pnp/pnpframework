using PnP.Framework.Migration.Diagnostics;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeTargetAdmission
    {
        public bool IsEligible { get; set; }

        public ContentTypeMaterializationDisposition Disposition { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
