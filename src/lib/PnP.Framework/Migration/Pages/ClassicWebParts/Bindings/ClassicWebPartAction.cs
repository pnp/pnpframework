using System;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    public enum ClassicWebPartDisposition
    {
        CopyCaptured = 1,
        RebindListAfterMaterialization = 2,
        Block = 3
    }

    public sealed class ClassicWebPartAction
    {
        public Guid SourceWebPartId { get; set; }

        public ClassicWebPartDisposition Disposition { get; set; }

        public Guid? SourceListWebId { get; set; }

        public Guid? SourceListId { get; set; }

        public Guid? SourceViewId { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetListServerRelativeUrl { get; set; }

        public string Reason { get; set; }
    }
}
