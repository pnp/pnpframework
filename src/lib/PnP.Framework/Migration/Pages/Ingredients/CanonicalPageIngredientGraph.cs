using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class CanonicalPageIngredientGraph
    {
        public string SchemaVersion { get; set; } = "pnp-page-ingredient-graph/v1";

        public IList<PageIngredientNode> Nodes { get; set; } = new List<PageIngredientNode>();

        public IList<PageIngredientEdge> Edges { get; set; } = new List<PageIngredientEdge>();
    }
}
