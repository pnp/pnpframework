using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class CanonicalPageIngredientGraph
    {
        public string SchemaVersion { get; set; } = "pnp-page-ingredient-graph/v1";

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string ProjectionVersion { get; set; }

        public IList<PageIngredientNode> Nodes { get; set; } = new List<PageIngredientNode>();

        public IList<PageIngredientEdge> Edges { get; set; } = new List<PageIngredientEdge>();
    }
}
