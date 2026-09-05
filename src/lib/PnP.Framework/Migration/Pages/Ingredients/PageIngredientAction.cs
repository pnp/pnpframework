using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class PageIngredientAction
    {
        public string ActionId { get; set; }

        public string IngredientId { get; set; }

        public IngredientCapability Capability { get; set; }

        public IngredientDisposition Disposition { get; set; }

        public string Realization { get; set; }

        public string TargetIdentity { get; set; }

        public string PolicyId { get; set; }

        public string PolicyVersion { get; set; }

        public string Reason { get; set; }

        public IList<string> ReleasedDependencyIngredientIds { get; set; } = new List<string>();

        public IList<string> VerificationAssertions { get; set; } = new List<string>();
    }
}
