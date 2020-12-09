using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown.Writers
{
    /// <summary>
    /// Class to write out the markdown for the base properties
    /// </summary>
    [TemplateSchemaWriter(WriterSequence = 1000,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateBasePropertiesWriter : PnPBaseSchemaWriter<ProvisioningTemplate>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            writer.WriteLine($"# Template - {template.Id}");
            writer.WriteLine();
            writer.WriteLine($"This is an export of the PnP Provisioning Template for this site.");
            writer.WriteLine();
        }
    }
}
