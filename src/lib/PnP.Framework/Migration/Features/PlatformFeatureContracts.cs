using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Features
{
    public enum PlatformFeatureScope
    {
        SiteCollection = 1,
        Web = 2
    }

    public enum PlatformFeatureMaterializationDisposition
    {
        EnsureActive = 1,
        Block = 2
    }

    public sealed class PlatformFeatureTargetProbe
    {
        public Guid FeatureId { get; set; }

        public PlatformFeatureScope Scope { get; set; }

        public string TargetWebUrl { get; set; }

        public bool TargetScopeExists { get; set; }

        public bool DeferredUntilTopologyMaterialization { get; set; }

        public bool IsActive { get; set; }

        public bool CanActivate { get; set; }

        public IList<string> AvailableContentTypeIds { get; set; } = new List<string>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsAdmitted => DeferredUntilTopologyMaterialization
            || (TargetScopeExists
                && (IsActive || CanActivate)
                && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker
                    && value.Severity != MigrationIssueSeverity.Error));
    }

    public sealed class PlatformFeatureMaterializationPlan
    {
        public Guid FeatureId { get; set; }

        public string Name { get; set; }

        public PlatformFeatureScope Scope { get; set; }

        public int DependencyOrder { get; set; }

        public IList<Guid> DependsOnFeatureIds { get; set; } = new List<Guid>();

        public IList<string> RequiredByContentTypeIds { get; set; } = new List<string>();

        public IList<string> ExpectedContentTypeIds { get; set; } = new List<string>();

        public string TargetWebUrl { get; set; }

        public PlatformFeatureMaterializationDisposition Disposition { get; set; }

        public string Reason { get; set; }

        public PlatformFeatureTargetProbe TargetProbe { get; set; }

        public bool IsExecutable => Disposition != PlatformFeatureMaterializationDisposition.Block
            && (TargetProbe == null || TargetProbe.IsAdmitted);
    }
}
