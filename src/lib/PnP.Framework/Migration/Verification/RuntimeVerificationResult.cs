namespace PnP.Framework.Migration.Verification
{
    public sealed class RuntimeVerificationResult
    {
        public string RequirementId { get; set; }

        public bool Passed { get; set; }

        public string EvidenceArtifactSha256 { get; set; }

        public string Message { get; set; }
    }
}
