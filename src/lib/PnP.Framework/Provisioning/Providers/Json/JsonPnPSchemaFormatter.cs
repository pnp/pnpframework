using Newtonsoft.Json;
using PnP.Framework.Provisioning.Providers.Xml;
using System;

namespace PnP.Framework.Provisioning.Providers.Json
{
    /// <summary>
    /// Helper class that abstracts from any specific version of JsonPnPSchemaFormatter
    /// </summary>
    public class JsonPnPSchemaFormatter : ITemplateFormatter
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        #region Static methods and properties

        /// <summary>
        /// Static method to retrieve an instance of the latest JsonPnPSchemaFormatter
        /// </summary>
        public static ITemplateFormatter LatestFormatter()
        {
            return LatestFormatter(null);
        }

        /// <summary>
        /// Static method to retrieve an instance of the latest JsonPnPSchemaFormatter
        /// </summary>
        /// <param name="serializerSettings">Custom <see cref="JsonSerializerSettings"/> to use</param>
        /// <returns></returns>
        public static ITemplateFormatter LatestFormatter(JsonSerializerSettings serializerSettings)
        {
            return new JsonPnPSchemaV202209Serializer(serializerSettings);
        }

        
        /// <summary>
        /// Static method to retrieve a specific JsonPnPSchemaFormatter instance
        /// </summary>
        /// <param name="version">Provisioning schema version</param>
        /// <returns></returns>
        public static ITemplateFormatter GetSpecificFormatter(XMLPnPSchemaVersion version)
        {
            return GetSpecificFormatter(version, null);
        }

        public static ITemplateFormatter GetSpecificFormatter(XMLPnPSchemaVersion version, JsonSerializerSettings serializerSettings)
        {
            switch (version)
            {
                case XMLPnPSchemaVersion.V202209:
                default:
                    return (new JsonPnPSchemaV202209Serializer(serializerSettings));
            }
        }
        /// <summary>
        /// Static method to retrieve a specific JsonPnPSchemaFormatter instance
        /// </summary>
        /// <param name="namespaceUri"></param>
        /// <returns></returns>
        public static ITemplateFormatter GetSpecificFormatter(string namespaceUri)
        {
            return GetSpecificFormatter(namespaceUri, null);
        }

        public static ITemplateFormatter GetSpecificFormatter(string namespaceUri, JsonSerializerSettings serializerSettings)
        {
            switch (namespaceUri)
            {
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2022_09:
                default:
                    return new JsonPnPSchemaV202209Serializer(serializerSettings);
            }
        }

        #endregion

        public bool IsValid(System.IO.Stream template)
        {
            return true;
        }

        public System.IO.Stream ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            String jsonString = JsonConvert.SerializeObject(template, new BasePermissionsConverter());
            Byte[] jsonBytes = System.Text.Encoding.Unicode.GetBytes(jsonString);
            System.IO.MemoryStream jsonStream = new System.IO.MemoryStream(jsonBytes)
            {
                Position = 0
            };

            return (jsonStream);
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template, string identifier)
        {
            using System.IO.StreamReader sr = new System.IO.StreamReader(template, System.Text.Encoding.Unicode);
            String jsonString = sr.ReadToEnd();

            Model.ProvisioningTemplate result = JsonConvert.DeserializeObject<Model.ProvisioningTemplate>(jsonString, new BasePermissionsConverter());
            return (result);
        }

        public System.IO.Stream ToFormattedHierarchy(Model.ProvisioningHierarchy hierarchy)
        {
            if (hierarchy == null)
            {
                throw new ArgumentNullException(nameof(hierarchy));
            }

            String jsonString = JsonConvert.SerializeObject(hierarchy, new BasePermissionsConverter());
            Byte[] jsonBytes = System.Text.Encoding.Unicode.GetBytes(jsonString);
            System.IO.MemoryStream jsonStream = new System.IO.MemoryStream(jsonBytes)
            {
                Position = 0
            };

            return (jsonStream);
        }

        public Model.ProvisioningHierarchy ToProvisioningHierarchy(System.IO.Stream hierarchy)
        {
            using System.IO.StreamReader sr = new System.IO.StreamReader(hierarchy, System.Text.Encoding.Unicode);
            String jsonString = sr.ReadToEnd();
            Model.ProvisioningHierarchy result = JsonConvert.DeserializeObject<Model.ProvisioningHierarchy>(jsonString, new BasePermissionsConverter());
            return (result);
        }

    }
}
