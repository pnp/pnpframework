using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeClosureNodePlan
    {
        public Guid SourceOwnerWebId { get; set; }

        public string SourceOwnerWebUrl { get; set; }

        public string TargetOwnerWebUrl { get; set; }

        public ContentTypeMaterializationPlan Schema { get; set; }

        public bool DeferredUntilTopologyMaterialization { get; set; }

        public ContentTypeTargetProbe TargetProbe { get; set; }

        public ContentTypeTargetAdmission TargetAdmission { get; set; }

        public string PlanDigest { get; set; }

        public bool IsExecutable => Schema != null
            && Schema.Disposition != ContentTypeMaterializationDisposition.Block
            && (TargetAdmission == null || TargetAdmission.IsEligible);
    }

    public sealed class ContentTypeClosurePlanBuildResult
    {
        public IList<ContentTypeClosureNodePlan> Nodes { get; set; } = new List<ContentTypeClosureNodePlan>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsExecutable => Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error)
            && Nodes.All(value => value.Schema != null && value.Schema.Disposition != ContentTypeMaterializationDisposition.Block);
    }
}
