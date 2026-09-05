using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class PageIngredientNode
    {
        public string Id { get; set; }

        public PageIngredientKind Kind { get; set; }

        public string Label { get; set; }

        public bool HasContent { get; set; }

        public PageIngredientOwnership Ownership { get; set; }

        public string SourceAuthority { get; set; }

        public string EvidenceDigest { get; set; }

        public string RuntimeRequirement { get; set; }

        public IList<string> EvidenceReferences { get; set; } = new List<string>();
    }
}
