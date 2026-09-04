using PnP.Framework.Migration.Taxonomy;
using System.Collections.Generic;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Taxonomy.Assets;

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

        public TaxonomyAssetMappingCatalog TaxonomyAssetMappingCatalog { get; set; }

        public TopologyPlanningPolicy TopologyPolicy { get; set; } = new TopologyPlanningPolicy();

        public IList<ListTargetOverride> ListTargetOverrides { get; set; } = new List<ListTargetOverride>();
    }
}
