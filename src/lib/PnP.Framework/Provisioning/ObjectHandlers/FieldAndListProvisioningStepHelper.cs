using System;
using System.Collections.Concurrent;
using System.Xml.Linq;
using Field = PnP.Framework.Provisioning.Model.Field;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    public static class FieldAndListProvisioningStepHelper
    {
        // Process wide cache, so it is shared by every provisioning run in the process.
        // A plain Dictionary corrupts itself when two runs provision fields at the same time.
        static readonly ConcurrentDictionary<Field, XElement> _fieldXmlDictionary = new ConcurrentDictionary<Field, XElement>();

        internal static Step GetFieldProvisioningStep(this Field templateField, TokenParser parser)
        {
            var schemaElement = GetCachedSchemaXml(templateField, parser);
            var type = (string)schemaElement.Attribute("Type");
            if (type != "Lookup" && type != "LookupMulti" && type != "Calculated")
            {
                return Step.ListAndStandardFields;
            }
            return Step.LookupFields;
        }

        internal static Guid GetFieldId(this Field templateField, TokenParser parser)
        {
            var schemaElement = GetCachedSchemaXml(templateField, parser);
            var id = (Guid)schemaElement.Attribute("ID");
            return id;
        }

        internal static XElement GetSchemaXml(this Field templateField, TokenParser parser, params string[] tokensToSkip)
        {
            return GetCachedSchemaXml(templateField, parser, tokensToSkip);
        }

        private static XElement GetCachedSchemaXml(Field templateField, TokenParser parser, params string[] tokensToSkip)
        {
            return _fieldXmlDictionary.GetOrAdd(
                templateField,
                field => XElement.Parse(parser.ParseXmlString(field.SchemaXml, tokensToSkip)));
        }

        public enum Step
        {
            /// <summary>
            /// The list itself and fields that aren't lookup fields are provisioned
            /// </summary>
            ListAndStandardFields,

            /// <summary>
            /// Focus on lookup fields. This assumes target lists are yet available
            /// </summary>
            LookupFields,

            /// <summary>
            /// Remaining list customization
            /// </summary>
            ListSettings,
            /// <summary>
            /// The handler is exporting
            /// </summary>
            Export
        }
    }
}