using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Fields;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal sealed class PublishingPageWriteResult
    {
        public List PagesLibrary { get; set; }

        public File TargetFile { get; set; }

        public ListItem TargetItem { get; set; }

        public IList<PageFieldImportResult> FieldResults { get; set; } = new List<PageFieldImportResult>();

        public bool ResumedExistingOwnedPage { get; set; }

        public bool PlannedFieldFailure => FieldResults.Any(result => result.Attempted && !result.Succeeded);
    }
}
