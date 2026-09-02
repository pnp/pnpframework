using PnP.Framework.Migration.Pages.Ingredients;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageIngredientGraphFactory
    {
        public static PageIngredientNode Node(
            string id,
            PageIngredientKind kind,
            string label,
            bool hasContent,
            PageIngredientOwnership ownership,
            string authority,
            string digest,
            string runtimeRequirement)
        {
            return new PageIngredientNode
            {
                Id = id,
                Kind = kind,
                Label = label,
                HasContent = hasContent,
                Ownership = ownership,
                SourceAuthority = authority,
                EvidenceDigest = digest,
                RuntimeRequirement = runtimeRequirement
            };
        }

        public static PageIngredientEdge Edge(
            string from,
            string to,
            PageIngredientRelationship relationship,
            PageIngredientRequirement requirement)
        {
            return new PageIngredientEdge
            {
                FromIngredientId = from,
                ToIngredientId = to,
                Relationship = relationship,
                Requirement = requirement
            };
        }
    }
}
