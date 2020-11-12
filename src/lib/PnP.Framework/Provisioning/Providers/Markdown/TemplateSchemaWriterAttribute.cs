using System;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    /// <summary>
    /// Attribute for Template Schema Serializers
    /// </summary>
    public class TemplateSchemaWriterAttribute : Attribute
    {
        /// <summary>
        /// The schemas supported by the serializer
        /// </summary>
        public PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaVersion MinimalSupportedSchemaVersion { get; set; }

        /// <summary>
        /// The sequence number for applying the writers
        /// </summary>
        /// <remarks>
        /// Should be a multiple of 100, to make room for future new insertions
        /// </remarks>
        public Int32 WriterSequence { get; set; } = 0;

        /// <summary>
        /// Defines the scope of the serializer
        /// </summary>
        /// <remarks>
        /// By default the serializers target a single Provisioning Template
        /// </remarks>
        public WriterScope Scope { get; set; } = WriterScope.ProvisioningTemplate;
    }

    /// <summary>
    /// Defines the scope of a serializer
    /// </summary>
    public enum WriterScope
    {
        /// <summary>
        /// The serializer targets a single Provisioning Template
        /// </summary>
        ProvisioningTemplate,
        /// <summary>
        /// The serializer targets a full Provisioning file but not a tenant Template
        /// </summary>
        Provisioning,
        /// <summary>
        /// The serializer targets a Provisioning Hierarchy
        /// </summary>
        ProvisioningHierarchy,
        /// <summary>
        /// The serializer targets the whole Tenant
        /// </summary>
        Tenant,
    }
}
