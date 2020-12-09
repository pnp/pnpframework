using PnP.Framework.Provisioning.Model;
using System;
using System.IO;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    /// <summary>
    /// Basic interface for every Schema Serializer type
    /// </summary>
    public interface IPnPSchemaWriter
    {
        /// <summary>
        /// Provides the name of the serializer type
        /// </summary>
        String Name { get; }

        /// <summary>
        /// The method to Writer a Domain Model object into an XML Schema based object 
        /// </summary>
        /// <param name="template">The PnP Provisioning Template object</param>
        /// <param name="persistence">The persistence layer object</param>
        /// <param name="writer"></param>
        void Writer(ProvisioningTemplate template, TextWriter writer);
    }
}
