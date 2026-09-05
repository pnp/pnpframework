using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Schema.ContentTypes;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutTargetAdmission
    {
        public bool IsEligible { get; set; }

        public PublishingPageLayoutMaterializationDisposition Disposition { get; set; }

        public ContentTypeTargetAdmission ContentTypeSchema { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
