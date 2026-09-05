namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class PageIngredientEdge
    {
        public string FromIngredientId { get; set; }

        public string ToIngredientId { get; set; }

        public PageIngredientRelationship Relationship { get; set; }

        public PageIngredientRequirement Requirement { get; set; }

        public string Condition { get; set; }
    }
}
