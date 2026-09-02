using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Schema.Fields
{
    internal static class FieldOwnershipClassifier
    {
        private static readonly HashSet<Guid> KnownTargetRuntimeFieldIds = new HashSet<Guid>
        {
            new Guid("ae3e2a36-125d-45d3-9051-744b513536a6"),
            new Guid("3b63724f-3418-461f-868b-7706f69b029c"),
            new Guid("c010d384-479c-494f-968c-c413dbe3de29"),
            new Guid("f3b0adf9-c1a2-4b02-920d-943fba4b3611"),
            new Guid("8f6b6dd8-9357-4019-8172-966fcd502ed2"),
            new Guid("e1a5b98c-dd71-426d-acb6-e478c7a5882f"),
            new Guid("f863c21f-5fdb-4a91-bb0c-5ae889190dd7")
        };

        public static FieldOwnership Classify(FieldSchemaSnapshot field, IReadOnlyCollection<FieldSchemaSnapshot> closure)
        {
            if (field == null)
            {
                throw new ArgumentNullException(nameof(field));
            }

            if (closure == null)
            {
                throw new ArgumentNullException(nameof(closure));
            }

            if (field.Role == FieldSchemaRole.InheritedFromParent || KnownTargetRuntimeFieldIds.Contains(field.Id))
            {
                return FieldOwnership.TargetRuntime;
            }

            if (field.Role == FieldSchemaRole.Dependency
                && field.Hidden
                && string.Equals(field.TypeAsString, "Note", StringComparison.OrdinalIgnoreCase)
                && closure.Any(candidate => candidate.Taxonomy?.HiddenTextFieldId == field.Id))
            {
                return FieldOwnership.GeneratedTaxonomyCompanion;
            }

            return HasPlatformSource(field.SchemaXml) ? FieldOwnership.TargetRuntime : FieldOwnership.UserDefined;
        }

        public static bool IsTargetRuntime(Guid fieldId, string schemaXml)
        {
            return KnownTargetRuntimeFieldIds.Contains(fieldId) || HasPlatformSource(schemaXml);
        }

        private static bool HasPlatformSource(string schemaXml)
        {
            if (string.IsNullOrWhiteSpace(schemaXml))
            {
                return false;
            }

            try
            {
                var source = XDocument.Parse(schemaXml).Root?.Attribute("SourceID")?.Value;
                return source?.StartsWith("http://schemas.microsoft.com/sharepoint/", StringComparison.OrdinalIgnoreCase) == true;
            }
            catch (System.Xml.XmlException)
            {
                return false;
            }
        }
    }
}
