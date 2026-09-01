namespace PnP.Framework.Migration.Diagnostics
{
    public sealed class MigrationIssue
    {
        public string Code { get; set; }

        public MigrationIssueSeverity Severity { get; set; }

        public string Subject { get; set; }

        public string Ingredient { get; set; }

        public string Message { get; set; }

        public string SourceIdentity { get; set; }

        public string TargetIdentity { get; set; }
    }
}
