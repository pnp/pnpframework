using PnP.Framework.Migration.Verification;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageRuntimeVerificationPolicy
    {
        private static readonly RuntimeVerificationRequirement[] Requirements =
        {
            Requirement("page-reachability", RuntimeVerificationRequirementKind.PageReachability, "Fresh navigation reaches the target classic page."),
            Requirement("error-shell-absence", RuntimeVerificationRequirementKind.ErrorShellAbsence, "The target is not a login, access denied, not found, or SharePoint error shell."),
            Requirement("authored-dom-equality", RuntimeVerificationRequirementKind.AuthoredDomEquality, "Normalized authored DOM is equal."),
            Requirement("resource-inventory-equality", RuntimeVerificationRequirementKind.ResourceInventoryEquality, "Authored resource inventory is equal."),
            Requirement("script-inventory-equality", RuntimeVerificationRequirementKind.ScriptInventoryEquality, "Authored script inventory is equal."),
            Requirement("inline-event-inventory-equality", RuntimeVerificationRequirementKind.InlineEventInventoryEquality, "Inline event inventory is equal."),
            Requirement("screenshot-capture", RuntimeVerificationRequirementKind.ScreenshotCapture, "Full-page and authored-canvas screenshots are captured.")
        };

        public static RuntimeVerificationManifest CreateManifest()
        {
            return new RuntimeVerificationManifest
            {
                Requirements = Requirements.Select(Copy).ToList()
            };
        }

        private static RuntimeVerificationRequirement Requirement(
            string id,
            RuntimeVerificationRequirementKind kind,
            string description)
        {
            return new RuntimeVerificationRequirement
            {
                Id = id,
                Kind = kind,
                Description = description,
                Required = true
            };
        }

        private static RuntimeVerificationRequirement Copy(RuntimeVerificationRequirement requirement)
        {
            return new RuntimeVerificationRequirement
            {
                Id = requirement.Id,
                Kind = requirement.Kind,
                Required = requirement.Required,
                Description = requirement.Description
            };
        }
    }
}
