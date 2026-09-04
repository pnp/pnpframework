using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageSnapshotAuthorizationEvidence
    {
        public static PageAssessmentEvidence Merge(
            PublishingPageCaptureBundle snapshot,
            PageAssessmentEvidence supplemental)
        {
            var source = snapshot?.Layout?.AuthorizationEvidence;
            if (source == null)
            {
                return supplemental;
            }

            LiteralHttpAuthorizationEvidence.Validate(source);
            var result = new PageAssessmentEvidence
            {
                AuthorizationFailures = (supplemental?.AuthorizationFailures
                        ?? Array.Empty<PageIngredientAuthorizationEvidence>())
                    .Where(value => value != null)
                    .Select(Copy)
                    .ToList()
            };
            Add(result.AuthorizationFailures, PublishingPageIngredientIds.Layout, source);
            Add(result.AuthorizationFailures, PublishingPageIngredientIds.ContentType, source);
            return result;
        }

        private static void Add(
            IList<PageIngredientAuthorizationEvidence> failures,
            string ingredientId,
            LiteralHttpAuthorizationEvidence source)
        {
            var existing = failures.SingleOrDefault(value =>
                string.Equals(value.IngredientId, ingredientId, StringComparison.Ordinal));
            if (existing != null)
            {
                if (existing.HttpStatusCode != source.HttpStatusCode
                    || !string.Equals(existing.RequestUri, source.RequestUri, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(existing.EvidenceSha256, source.EvidenceSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        $"Conflicting authorization evidence exists for ingredient '{ingredientId}'.");
                }
                return;
            }

            failures.Add(new PageIngredientAuthorizationEvidence
            {
                IngredientId = ingredientId,
                Operation = source.Operation,
                RequestUri = source.RequestUri,
                HttpStatusCode = source.HttpStatusCode,
                ObservedAtUtc = source.ObservedAtUtc,
                EvidenceSource = "snapshot.layout.authorizationEvidence",
                EvidenceSha256 = source.EvidenceSha256
            });
        }

        private static PageIngredientAuthorizationEvidence Copy(
            PageIngredientAuthorizationEvidence source)
        {
            return new PageIngredientAuthorizationEvidence
            {
                IngredientId = source.IngredientId,
                Operation = source.Operation,
                RequestUri = source.RequestUri,
                HttpStatusCode = source.HttpStatusCode,
                ObservedAtUtc = source.ObservedAtUtc,
                EvidenceSource = source.EvidenceSource,
                EvidenceSha256 = source.EvidenceSha256
            };
        }
    }
}
