using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    /// <summary>
    /// Establishes the only final ingredient-action boundary that may emit Block:
    /// retained literal wire HTTP 401/403 evidence. Domain planners may use their
    /// own local Block disposition, but that is normalized to Defer before this
    /// policy is applied.
    /// </summary>
    internal static class PublishingPageIngredientAuthorizationPolicy
    {
        public static IReadOnlyDictionary<string, LiteralHttpAuthorizationEvidence> GetEvidence(
            PublishingPageCaptureBundle snapshot)
        {
            var source = snapshot?.Layout?.AuthorizationEvidence;
            if (source == null)
            {
                return new Dictionary<string, LiteralHttpAuthorizationEvidence>(StringComparer.Ordinal);
            }

            LiteralHttpAuthorizationEvidence.Validate(source);
            return new Dictionary<string, LiteralHttpAuthorizationEvidence>(StringComparer.Ordinal)
            {
                [PublishingPageIngredientIds.Layout] = source,
                [PublishingPageIngredientIds.ContentType] = source
            };
        }

        public static void Apply(
            PublishingPageCaptureBundle snapshot,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var pair in GetEvidence(snapshot))
            {
                if (actions == null || !actions.TryGetValue(pair.Key, out var action) || action == null)
                {
                    throw new InvalidDataException(
                        $"Literal authorization evidence references unknown ingredient '{pair.Key}'.");
                }

                var evidence = pair.Value;
                action.Capability = IngredientCapability.Missing;
                action.Disposition = IngredientDisposition.Block;
                action.Realization = "none";
                action.PolicyId = "policy.authorization.literal-http";
                action.Reason = $"Source ingredient request '{evidence.Operation}' returned literal HTTP {evidence.HttpStatusCode}.";
                action.VerificationAssertions = (action.VerificationAssertions ?? new List<string>())
                    .Concat(new[]
                    {
                        $"Authorization evidence has SHA-256 '{evidence.EvidenceSha256}'."
                    })
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
            }
        }
    }
}
