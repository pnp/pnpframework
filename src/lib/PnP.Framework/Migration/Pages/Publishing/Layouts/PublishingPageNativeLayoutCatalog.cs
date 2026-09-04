using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal sealed class PublishingPageNativeLayoutProfile
    {
        public string FileName { get; set; }

        public string Title { get; set; }

        public string AssociatedContentTypeName { get; set; }

        public string AssociatedContentTypeId { get; set; }
    }

    internal static class PublishingPageNativeLayoutCatalog
    {
        private static readonly IReadOnlyDictionary<string, PublishingPageNativeLayoutProfile> Profiles =
            new Dictionary<string, PublishingPageNativeLayoutProfile>(StringComparer.OrdinalIgnoreCase)
            {
                ["EnterpriseWiki.aspx"] = new PublishingPageNativeLayoutProfile
                {
                    FileName = "EnterpriseWiki.aspx",
                    Title = "Basic Page",
                    AssociatedContentTypeName = "Enterprise Wiki Page",
                    AssociatedContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage
                },
                ["BlankWebPartPage.aspx"] = new PublishingPageNativeLayoutProfile
                {
                    FileName = "BlankWebPartPage.aspx",
                    Title = "Blank Web Part page",
                    AssociatedContentTypeName = "Welcome Page",
                    AssociatedContentTypeId = BuiltInContentTypeId.WelcomePage
                }
            };

        public static bool TryGetUnavailableSourceSubstitution(
            PublishingPageLayoutSnapshot layout,
            string fileName,
            out PublishingPageNativeLayoutProfile profile)
        {
            profile = null;
            if (layout == null
                || layout.EvidenceState == PublishingPageLayoutEvidenceState.Readable
                || layout.Availability != EvidenceAvailability.Unavailable
                || string.IsNullOrWhiteSpace(fileName)
                || !Profiles.TryGetValue(fileName, out var candidate)
                || !string.Equals(layout.Description?.Trim(), candidate.Title, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            profile = candidate;
            return true;
        }
    }
}
