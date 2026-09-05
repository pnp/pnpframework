using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Lists.Fields
{
    internal static class ListFieldSnapshotReader
    {
        public static IList<ListFieldSnapshot> Read(ClientContext context, FieldCollection fields)
        {
            var taxonomyFields = new Dictionary<Guid, TaxonomyField>();
            foreach (var field in fields.Where(value => value.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase)))
            {
                var taxonomy = context.CastTo<TaxonomyField>(field);
                context.Load(
                    taxonomy,
                    value => value.SspId,
                    value => value.TermSetId,
                    value => value.AnchorId,
                    value => value.TextField,
                    value => value.Open);
                taxonomyFields[field.Id] = taxonomy;
            }
            if (taxonomyFields.Count > 0)
            {
                context.ExecuteQueryRetry();
            }

            return fields.Select(field => Create(field, taxonomyFields)).OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static ListFieldSnapshot Create(Field field, IDictionary<Guid, TaxonomyField> taxonomyFields)
        {
            Guid? lookupWebId = null;
            Guid? lookupListId = null;
            string lookupField = null;
            try
            {
                var root = XDocument.Parse(field.SchemaXml).Root;
                lookupWebId = ParseGuid(root == null ? null : (string)root.Attribute("WebId"));
                lookupListId = ParseGuid(root == null ? null : (string)root.Attribute("List"));
                lookupField = root == null ? null : (string)root.Attribute("ShowField");
            }
            catch (System.Xml.XmlException)
            {
            }

            TaxonomyField taxonomy;
            return new ListFieldSnapshot
            {
                Id = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                Group = field.Group,
                SchemaXml = field.SchemaXml,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(field.SchemaXml ?? string.Empty),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml),
                Hidden = field.Hidden,
                ReadOnly = field.ReadOnlyField,
                Required = field.Required,
                FromBaseType = field.FromBaseType,
                Sealed = field.Sealed,
                SourceLookupWebId = lookupWebId,
                SourceLookupListId = lookupListId,
                LookupField = lookupField,
                Taxonomy = taxonomyFields.TryGetValue(field.Id, out taxonomy)
                    ? new TaxonomyFieldBindingSnapshot
                    {
                        SourceTermStoreId = taxonomy.SspId,
                        SourceTermSetId = taxonomy.TermSetId,
                        AnchorTermId = taxonomy.AnchorId,
                        HiddenTextFieldId = taxonomy.TextField,
                        Open = taxonomy.Open
                    }
                    : null
            };
        }

        private static Guid? ParseGuid(string value)
        {
            Guid result;
            return Guid.TryParse((value ?? string.Empty).Trim().Trim('{', '}'), out result) && result != Guid.Empty ? result : (Guid?)null;
        }
    }
}
