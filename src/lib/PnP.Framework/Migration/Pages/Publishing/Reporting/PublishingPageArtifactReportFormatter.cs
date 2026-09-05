using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPageArtifactReportFormatter
    {
        public static string Artifact(ArtifactReference artifact)
        {
            if (artifact == null)
            {
                return null;
            }

            return $"sha256={Format(artifact.Sha256)}; length={artifact.Length}; mediaType={Format(artifact.MediaType)}; "
                + $"contentEncoding={Format(artifact.ContentEncoding)}; originalName={Format(artifact.OriginalName)}; "
                + $"availability={artifact.Availability}; lineage={Lineage(artifact.Lineage)}";
        }

        public static string Sources(IEnumerable<EvidenceSource> sources)
        {
            return Join((sources ?? Array.Empty<EvidenceSource>()).Select(item =>
                $"exchangeId={Format(item.ExchangeId)}; payloadSha256={Format(item.PayloadSha256)}; selector={Format(item.Selector)}"));
        }

        public static string Taxonomy(TaxonomyFieldBindingSnapshot taxonomy)
        {
            return taxonomy == null
                ? null
                : $"sourceTermStoreId={taxonomy.SourceTermStoreId:D}; sourceTermSetId={taxonomy.SourceTermSetId:D}; "
                    + $"hiddenTextFieldId={taxonomy.HiddenTextFieldId:D}; open={taxonomy.Open}";
        }

        private static string Lineage(ArtifactLineage lineage)
        {
            if (lineage == null)
            {
                return null;
            }

            return $"inputExchangeIds=[{Format(Join(lineage.InputExchangeIds))}]; "
                + $"inputPayloadSha256=[{Format(Join(lineage.InputPayloadSha256))}]; "
                + $"projectorId={Format(lineage.ProjectorId)}; projectorVersion={Format(lineage.ProjectorVersion)}; "
                + $"outputSchemaVersion={Format(lineage.OutputSchemaVersion)}; outputSha256={Format(lineage.OutputSha256)}";
        }

        private static string Format(object value) => PublishingPageReportValueFormatter.Format(value);

        private static string Join(IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);
    }
}
