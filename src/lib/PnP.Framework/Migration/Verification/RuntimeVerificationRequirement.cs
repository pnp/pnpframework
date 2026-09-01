namespace PnP.Framework.Migration.Verification
{
    public sealed class RuntimeVerificationRequirement
    {
        public string Id { get; set; }

        public RuntimeVerificationRequirementKind Kind { get; set; }

        public bool Required { get; set; } = true;

        public string Description { get; set; }
    }
}
