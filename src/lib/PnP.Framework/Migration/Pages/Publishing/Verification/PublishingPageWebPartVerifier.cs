using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageWebPartVerifier
    {
        public static IList<PublishingPageWebPartVerificationResult> Verify(
            ClientContext context,
            string pagePath,
            IEnumerable<ClassicWebPartSnapshot> expectedWebParts,
            IEnumerable<ClassicListWebPartBindingSnapshot> bindings,
            IEnumerable<ClassicWebPartAction> actions,
            IEnumerable<ListMaterializationReceipt> listReceipts,
            IEnumerable<PageTextReplacement> replacements,
            bool exactInventory = true)
        {
            var actual = Read(context, pagePath);
            var unused = actual.ToList();
            var results = new List<PublishingPageWebPartVerificationResult>();
            var actionByWebPart = actions.ToDictionary(value => value.SourceWebPartId);
            var bindingByWebPart = bindings.ToDictionary(value => value.SourceWebPartId);
            var receiptByList = listReceipts.ToDictionary(value => value.SourceListId);
            foreach (var expected in expectedWebParts
                         .OrderBy(item => item.ZoneId, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(item => item.ZoneIndex)
                         .ThenBy(item => item.Id))
            {
                ClassicWebPartAction action;
                ClassicListWebPartBindingSnapshot binding;
                ListMaterializationReceipt listReceipt = null;
                actionByWebPart.TryGetValue(expected.Id, out action);
                bindingByWebPart.TryGetValue(expected.Id, out binding);
                if (binding != null)
                {
                    receiptByList.TryGetValue(binding.SourceListId, out listReceipt);
                }
                var expectedXml = ClassicWebPartReplayComposer.Compose(
                    expected,
                    action,
                    binding,
                    listReceipt,
                    pagePath,
                    replacements);
                var expectedDigest = PublishingPageDigest.ComputeSha256(expectedXml);
                var match = unused.FirstOrDefault(item =>
                    string.Equals(item.ExportSha256, expectedDigest, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(item.ZoneId, expected.ZoneId, StringComparison.OrdinalIgnoreCase)
                    && item.ZoneIndex == expected.ZoneIndex);
                var storageCanonicalMatch = false;
                if (match == null && action?.Disposition == ClassicWebPartDisposition.RebindListAfterMaterialization)
                {
                    var expectedCanonical = ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(expectedXml);
                    match = unused.FirstOrDefault(item =>
                        string.Equals(item.ZoneId, expected.ZoneId, StringComparison.OrdinalIgnoreCase)
                        && item.ZoneIndex == expected.ZoneIndex
                        && string.Equals(
                            ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(item.ExportXml),
                            expectedCanonical,
                            StringComparison.Ordinal));
                    storageCanonicalMatch = match != null;
                }
                if (match == null)
                {
                    match = unused.FirstOrDefault(item =>
                        string.Equals(item.ExportSha256, expectedDigest, StringComparison.OrdinalIgnoreCase));
                }

                var passed = match != null
                    && string.Equals(match.ZoneId, expected.ZoneId, StringComparison.OrdinalIgnoreCase)
                    && match.ZoneIndex == expected.ZoneIndex
                    && match.Hidden == expected.Hidden;
                results.Add(new PublishingPageWebPartVerificationResult
                {
                    SourceWebPartId = expected.Id,
                    TargetWebPartId = match?.Id,
                    ExpectedZoneId = expected.ZoneId,
                    ExpectedZoneIndex = expected.ZoneIndex,
                    ActualZoneId = match?.ZoneId,
                    ActualZoneIndex = match?.ZoneIndex,
                    ExpectedExportSha256 = expectedDigest,
                    ActualExportSha256 = match?.ExportSha256,
                    Passed = passed,
                    Message = Describe(expected, match, passed, storageCanonicalMatch)
                });
                if (match != null)
                {
                    unused.Remove(match);
                }
            }

            foreach (var extra in exactInventory ? unused : Enumerable.Empty<ClassicWebPartSnapshot>())
            {
                results.Add(new PublishingPageWebPartVerificationResult
                {
                    TargetWebPartId = extra.Id,
                    ActualZoneId = extra.ZoneId,
                    ActualZoneIndex = extra.ZoneIndex,
                    ActualExportSha256 = extra.ExportSha256,
                    Passed = false,
                    Message = "The target contains an unplanned shared Web Part."
                });
            }

            return results;
        }

        private static IList<ClassicWebPartSnapshot> Read(ClientContext context, string pagePath)
        {
            var web = context.Web;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var manager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var definitions = web.GetWebParts(pagePath).ToArray();
            var result = new List<ClassicWebPartSnapshot>();
            foreach (var definition in definitions)
            {
                var export = manager.ExportWebPart(definition.Id);
                context.ExecuteQueryRetry();
                var xml = export.Value ?? string.Empty;
                result.Add(new ClassicWebPartSnapshot
                {
                    Id = definition.Id,
                    Title = definition.WebPart.Title,
                    ZoneId = definition.ZoneId,
                    ZoneIndex = definition.WebPart.ZoneIndex,
                    Hidden = definition.WebPart.Hidden,
                    ExportXml = xml,
                    ExportSha256 = PublishingPageDigest.ComputeSha256(xml)
                });
            }

            return result;
        }

        private static string Describe(
            ClassicWebPartSnapshot expected,
            ClassicWebPartSnapshot actual,
            bool passed,
            bool storageCanonicalMatch)
        {
            if (actual == null)
            {
                return "No target Web Part has the approved export digest.";
            }

            if (passed)
            {
                return storageCanonicalMatch
                    ? "Storage-canonical list binding, zone placement, and hidden state match; SharePoint regenerated only runtime View identity or equivalent empty/null XML representation."
                    : "Export digest, zone placement, and hidden state match.";
            }

            return $"Expected zone '{expected.ZoneId}' index {expected.ZoneIndex} hidden={expected.Hidden}; actual zone '{actual.ZoneId}' index {actual.ZoneIndex} hidden={actual.Hidden}.";
        }
    }
}
