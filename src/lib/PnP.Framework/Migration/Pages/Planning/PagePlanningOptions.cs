using PnP.Framework.Migration.Taxonomy;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Planning
{
    public sealed class PagePlanningOptions
    {
        public string TargetPageServerRelativeUrl { get; set; }

        public bool RequireInheritedPermissions { get; set; } = true;

        public bool BlockOnManagedMetadata { get; set; } = true;

        public bool AllowExternalResourceReferences { get; set; } = true;

        public bool CreateOnly { get; set; } = true;

        public IList<TaxonomyTargetMapping> TaxonomySchemaMappings { get; set; } = new List<TaxonomyTargetMapping>();
    }
}
