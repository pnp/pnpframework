using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageIngredientActionFactory
    {
        public static PageIngredientAction Create(
            string ingredientId,
            IngredientCapability capability,
            IngredientDisposition disposition,
            string realization,
            string policyId,
            string reason,
            string targetIdentity = null,
            params string[] verificationAssertions)
        {
            return new PageIngredientAction
            {
                ActionId = "action:" + ingredientId,
                IngredientId = ingredientId,
                Capability = capability,
                Disposition = disposition,
                Realization = realization,
                TargetIdentity = targetIdentity,
                PolicyId = policyId,
                PolicyVersion = "1",
                Reason = reason,
                VerificationAssertions = (verificationAssertions ?? Array.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .ToList()
            };
        }

        public static void Add(
            IDictionary<string, PageIngredientAction> actions,
            PageIngredientAction action)
        {
            if (action != null && !actions.ContainsKey(action.IngredientId))
            {
                actions.Add(action.IngredientId, action);
            }
        }
    }
}
