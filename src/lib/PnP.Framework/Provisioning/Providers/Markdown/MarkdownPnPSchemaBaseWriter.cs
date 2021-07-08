using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    /// <summary>
    /// Base class for writing of provisioning markdown
    /// with the new schema serializer
    /// </summary>
    internal abstract class MarkdownPnPSchemaBaseWriter
    {
        private TemplateProviderBase _provider;
        private readonly Stream _referenceSchema;

        protected TemplateProviderBase Provider => _provider;

        public MarkdownPnPSchemaBaseWriter(Stream referenceSchema)
        {
            this._referenceSchema = referenceSchema ??
                throw new ArgumentNullException(nameof(referenceSchema));
        }

        public abstract string NamespacePrefix { get; }
        public abstract string NamespaceUri { get; }

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        /*
        private static PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion GetCurrentSchemaVersion()
        {
            var currentSchemaTemplateNamespace = typeof(PnP.Framework.Provisioning.Providers.Xml.TSchemaTemplate).Namespace;
            var currentSchemaVersionString = $"V{currentSchemaTemplateNamespace.Substring(currentSchemaTemplateNamespace.IndexOf(".Xml.") + 6)}";
            var currentSchemaVersion = (PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion)Enum.Parse(typeof(PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion), currentSchemaVersionString);
            return currentSchemaVersion;
        }*/

        protected virtual void WriteTemplate(TextWriter persistenceTemplate, ProvisioningTemplate template)
        {
            // Get all ProvisioningTemplate-level serializers to run in automated mode, ordered by DeserializationSequence
            var serializers = GetSerializersForCurrentContext(WriterScope.ProvisioningTemplate, a => a?.WriterSequence);

            // Invoke all the ProvisioningTemplate-level serializers
            InvokeSerializers(template, persistenceTemplate, serializers);
        }

        private IOrderedEnumerable<IGrouping<string, Type>> GetSerializersForCurrentContext(WriterScope scope,
            Func<TemplateSchemaWriterAttribute, Int32?> sortingSelector)
        {
            // Get all serializers to run in automated mode, ordered by sortingSelector
            var currentAssembly = this.GetType().Assembly;

            //PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion currentSchemaVersion = GetCurrentSchemaVersion();

            var serializers = currentAssembly.GetTypes()
                // Get all the serializers
                .Where(t => t.GetInterface(typeof(IPnPSchemaWriter).FullName) != null
                       && t.BaseType.Name == typeof(Xml.PnPBaseSchemaSerializer<>).Name)
                // Filter out those that are not targeting the current schema version or that are not in scope Template
                .Where(t =>
                {
                    var a = t.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                    return (a.Scope == scope);
                })
                // Order the remainings by supported schema version descendant, to get first the newest ones
                .OrderByDescending(s =>
                {
                    var a = s.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion);
                }
                )
                // Group those with the same target type (which is the first generic Type argument)
                .GroupBy(t => t.BaseType.GenericTypeArguments.FirstOrDefault()?.FullName)
                // Order the result by SerializationSequence
                .OrderBy(g =>
                {
                    var maxInGroup = g.OrderByDescending(s =>
                    {
                        var a = s.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault();
                        return (a.MinimalSupportedSchemaVersion);
                    }
                    ).FirstOrDefault();
                    return sortingSelector(maxInGroup.GetCustomAttributes<TemplateSchemaWriterAttribute>(false).FirstOrDefault());
                });
            return serializers;
        }


        private static void InvokeSerializers(ProvisioningTemplate template, TextWriter writer,
            IOrderedEnumerable<IGrouping<string, Type>> serializers)
        {
            foreach (var group in serializers)
            {
                // Get the first serializer only for each group (i.e. the most recent one for the current schema)
                var serializerType = group.FirstOrDefault();
                if (serializerType != null)
                {
                    // Create an instance of the serializer
                    var serializer = Activator.CreateInstance(serializerType) as IPnPSchemaWriter;
                    if (serializer != null)
                    {
                            serializer.Writer(template, writer);
                    }
                }
            }
        }
    }

}
